import { promises as fs } from "node:fs";
import { basename, extname } from "node:path";

import type {
  AnyFileSource,
  FileHandler,
  FileHandlerResult,
  FileMetadata,
} from "@anyfile/core";
import * as XLSX from "xlsx";

import type {
  ExcelCellValue,
  ExcelFileData,
  ExcelMetadata,
  ExcelReadOptions,
  ExcelWorksheetDescriptor,
} from "./types";

const EXCEL_EXTENSIONS = ["xls", "xlsx", "xlsm", "xlsb"];
const XLS_SIGNATURE = Buffer.from([0xd0, 0xcf, 0x11, 0xe0]);
const XLSX_SIGNATURE = Buffer.from([0x50, 0x4b, 0x03, 0x04]);

export function createExcelHandler(): FileHandler<ExcelFileData> {
  return {
    type: "excel",
    extensions: EXCEL_EXTENSIONS,
    detect: async (source) => detectExcelSource(source),
    open: async ({ source, metadata }) => {
      const payload = await loadSource(source);
      const workbook = XLSX.read(payload.buffer, { type: "buffer" });

      const fileMetadata: FileMetadata = {
        name: metadata?.name ?? payload.name,
        size: metadata?.size ?? payload.size,
        type: "excel",
        createdAt: metadata?.createdAt ?? payload.createdAt,
        modifiedAt: metadata?.modifiedAt ?? payload.modifiedAt,
      };

      return buildHandlerResult(workbook, fileMetadata);
    },
  };
}

async function detectExcelSource(source: AnyFileSource): Promise<boolean> {
  if (typeof source === "string") {
    return isExcelExtension(source);
  }

  const buffer = toBuffer(source);
  if (buffer.length < 4) {
    return false;
  }

  const signature = buffer.subarray(0, 4);
  return signature.equals(XLS_SIGNATURE) || signature.equals(XLSX_SIGNATURE);
}

async function loadSource(source: AnyFileSource) {
  if (typeof source === "string") {
    const [buffer, stats] = await Promise.all([
      fs.readFile(source),
      fs.stat(source),
    ]);

    return {
      buffer,
      name: basename(source),
      size: stats.size,
      createdAt: stats.birthtime,
      modifiedAt: stats.mtime,
    };
  }

  const buffer = toBuffer(source);
  return {
    buffer,
    name: "buffer.xlsx",
    size: buffer.byteLength,
    createdAt: undefined,
    modifiedAt: undefined,
  };
}

function buildHandlerResult(
  workbook: XLSX.WorkBook,
  metadata: FileMetadata
): FileHandlerResult<ExcelFileData> {
  const createData = () => createExcelFileData(workbook);

  return {
    type: "excel",
    metadata,
    read: async () => createData(),
    write: async (outputPath, data) => {
      const payload = data ?? createData();
      const workbookToWrite = payload.workbook ?? workbook;
      const bookType = resolveBookType(outputPath);
      const buffer = XLSX.write(workbookToWrite, {
        bookType,
        type: "buffer",
      });
      await fs.writeFile(outputPath, buffer);
    },
    convert: async (toType) => {
      if (toType !== "csv") {
        throw new Error(`Conversion from Excel to "${toType}" is not implemented yet.`);
      }

      const csv = workbookToCSV(workbook);
      const csvMetadata: FileMetadata = {
        ...metadata,
        type: "csv",
        name: replaceExtension(metadata.name, "csv"),
      };

      return {
        type: "csv",
        metadata: csvMetadata,
        read: async () => csv,
        write: async (output, data) => {
          const content = data ?? csv;
          await fs.writeFile(output, content, "utf8");
        },
      };
    },
  };
}

function createExcelFileData(workbook: XLSX.WorkBook): ExcelFileData {
  const describeSheets = (): ExcelWorksheetDescriptor[] =>
    workbook.SheetNames.map((sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      const rangeRef = sheet?.["!ref"];
      const range = rangeRef ? XLSX.utils.decode_range(rangeRef) : undefined;

      return {
        name: sheetName,
        range: rangeRef,
        rowCount: range ? range.e.r - range.s.r + 1 : 0,
        columnCount: range ? range.e.c - range.s.c + 1 : 0,
      };
    });

  const readSheet = async (
    nameOrIndex: string | number = 0,
    options?: ExcelReadOptions
  ) => {
    const sheetName = resolveSheetName(workbook, nameOrIndex);
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found.`);
    }

    const range = options?.range ?? sheet["!ref"];
    const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
      header: 1,
      blankrows: false,
      defval: null,
      range,
    }) as unknown as unknown[][];

    const headerRowIndex = Math.max(0, (options?.headerRow ?? 1) - 1);
    const headerRow = (rows[headerRowIndex] ?? []) as unknown[];
    const dataRows = rows.slice(headerRowIndex + 1);

    return dataRows.map((row) => {
      const record: Record<string, unknown> = {};
      headerRow.forEach((header, index) => {
        const key =
          header !== undefined && header !== null && `${header}`.trim().length
            ? `${header}`
            : `column_${index + 1}`;
        record[key] = Array.isArray(row) ? row[index] ?? null : null;
      });
      return record;
    });
  };

  const toJSON = async () => {
    const entries = await Promise.all(
      describeSheets().map(async (descriptor) => [
        descriptor.name,
        await readSheet(descriptor.name),
      ])
    );

    return Object.fromEntries(entries);
  };

  const api: ExcelFileData = {
    workbook,
    getSheets: describeSheets,
    getSheetNames: () => workbook.SheetNames.slice(),
    readSheet,
    getCell: (sheet, row, column) =>
      getCellValue(workbook, sheet, row, column),
    setCell: (sheet, row, column, value) =>
      setCellValue(workbook, sheet, row, column, value),
    addSheet: (name: string) => addWorksheet(workbook, name),
    addSheetFromCSV: (name: string, csv: string) =>
      addWorksheetFromCSV(workbook, name, csv),
    deleteSheet: (name: string) => deleteWorksheet(workbook, name),
    getMetadata: () => extractMetadata(workbook),
    toCSV: (sheet?: string | number) => worksheetToCSV(workbook, sheet),
    toJSON,
    worksheets: describeSheets(),
  };

  Object.defineProperty(api, "worksheets", {
    get: describeSheets,
    enumerable: true,
  });

  return api;
}

function resolveSheetName(workbook: XLSX.WorkBook, input: string | number) {
  if (typeof input === "string") {
    return input;
  }
  const index = input ?? 0;
  return workbook.SheetNames[index] ?? workbook.SheetNames[0];
}

function toBuffer(source: ArrayBuffer | ArrayBufferView): Buffer {
  if (source instanceof ArrayBuffer) {
    return Buffer.from(source);
  }

  if (ArrayBuffer.isView(source)) {
    return Buffer.from(
      source.buffer,
      source.byteOffset,
      source.byteLength
    );
  }

  throw new Error("Unsupported source provided to Excel handler.");
}

function isExcelExtension(path: string): boolean {
  const extension = extname(path).replace(".", "").toLowerCase();
  return EXCEL_EXTENSIONS.includes(extension);
}

function resolveBookType(outputPath: string) {
  const extension = extname(outputPath).replace(".", "").toLowerCase();
  switch (extension) {
    case "xlsm":
      return "xlsm";
    case "xlsb":
      return "xlsb";
    case "xls":
      return "xls";
    default:
      return "xlsx";
  }
}

function addWorksheet(workbook: XLSX.WorkBook, name: string) {
  const sheetName = name.trim();
  if (!sheetName) {
    throw new Error("Sheet name cannot be empty.");
  }

  if (workbook.Sheets[sheetName]) {
    throw new Error(`Sheet "${sheetName}" already exists.`);
  }

  const worksheet = XLSX.utils.aoa_to_sheet([]);
  XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
}

function addWorksheetFromCSV(
  workbook: XLSX.WorkBook,
  name: string,
  csvText: string
) {
  const sheetName = name.trim();
  if (!sheetName) {
    throw new Error("Sheet name cannot be empty.");
  }

  if (workbook.Sheets[sheetName]) {
    throw new Error(`Sheet "${sheetName}" already exists.`);
  }

  const csvWorkbook = XLSX.read(csvText, { type: "string" });
  const worksheet = csvWorkbook.Sheets[csvWorkbook.SheetNames[0]];
  if (!worksheet) {
    throw new Error("Unable to parse CSV content.");
  }

  XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
}

function deleteWorksheet(workbook: XLSX.WorkBook, name: string) {
  if (!workbook.Sheets[name]) {
    throw new Error(`Sheet "${name}" not found.`);
  }

  delete workbook.Sheets[name];
  const index = workbook.SheetNames.indexOf(name);
  if (index >= 0) {
    workbook.SheetNames.splice(index, 1);
  }
}

function getCellValue(
  workbook: XLSX.WorkBook,
  sheet: string | number,
  row: number,
  column: number
) {
  const sheetName = resolveSheetName(workbook, sheet);
  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) {
    throw new Error(`Sheet "${sheetName}" not found.`);
  }

  const cellAddress = XLSX.utils.encode_cell({ r: row - 1, c: column - 1 });
  const cell = worksheet[cellAddress];
  if (!cell) {
    return undefined;
  }

  return {
    address: cellAddress,
    value: cell.v ?? null,
    raw: cell.w,
    type: cell.t,
  };
}

function setCellValue(
  workbook: XLSX.WorkBook,
  sheet: string | number,
  row: number,
  column: number,
  value: ExcelCellValue
) {
  const sheetName = resolveSheetName(workbook, sheet);
  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) {
    throw new Error(`Sheet "${sheetName}" not found.`);
  }

  if (row < 1 || column < 1) {
    throw new Error("Row and column must be 1-based positive integers.");
  }

  XLSX.utils.sheet_add_aoa(
    worksheet,
    [[value ?? null]],
    {
      origin: { r: row - 1, c: column - 1 },
    }
  );
}

function extractMetadata(workbook: XLSX.WorkBook): ExcelMetadata {
  const props = workbook.Props ?? {};
  return {
    title: props.Title ?? undefined,
    subject: props.Subject ?? undefined,
    author: props.Author ?? undefined,
    manager: props.Manager ?? undefined,
    company: props.Company ?? undefined,
    category: props.Category ?? undefined,
    keywords: props.Keywords ?? undefined,
    comments: props.Comments ?? undefined,
    lastAuthor: props.LastAuthor ?? undefined,
    sheetCount: workbook.SheetNames.length,
    createdAt: props.CreatedDate ?? undefined,
    modifiedAt: props.ModifiedDate ?? undefined,
  };
}

function worksheetToCSV(
  workbook: XLSX.WorkBook,
  sheet: string | number = 0
): string {
  const sheetName = resolveSheetName(workbook, sheet);
  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) {
    throw new Error(`Sheet "${sheetName}" not found.`);
  }
  return XLSX.utils.sheet_to_csv(worksheet);
}

function workbookToCSV(workbook: XLSX.WorkBook): string {
  if (workbook.SheetNames.length === 0) {
    return "";
  }

  // Concatenate sheets separated by blank line and sheet title.
  return workbook.SheetNames.map((name) => {
    const csv = worksheetToCSV(workbook, name);
    return `# ${name}\n${csv}`.trim();
  }).join("\n\n");
}

function replaceExtension(filename: string, newExt: string) {
  if (!filename) {
    return `workbook.${newExt}`;
  }

  const base = filename.replace(/\.[^.]+$/, "");
  return `${base}.${newExt}`;
}

