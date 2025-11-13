import { promises as fs } from "node:fs";
import { basename, extname } from "node:path";

import type {
  AnyFileSource,
  FileHandler,
  FileHandlerResult,
  FileMetadata,
} from "@anyfile/core";
import * as XLSX from "xlsx";
import XLSX_CALC = require("xlsx-calc");

import type {
  ExcelCellStyle,
  ExcelCellValue,
  ExcelCircularReference,
  ExcelFileData,
  ExcelMetadata,
  ExcelReadOptions,
  ExcelSetCellOptions,
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
      const workbook = XLSX.read(payload.buffer, {
        type: "buffer",
        cellFormula: true,
      });

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
    setCell: (sheet, row, column, value, options) =>
      setCellValue(workbook, sheet, row, column, value, options),
    addSheet: (name: string) => addWorksheet(workbook, name),
    addSheetFromCSV: (name: string, csv: string) =>
      addWorksheetFromCSV(workbook, name, csv),
    deleteSheet: (name: string) => deleteWorksheet(workbook, name),
    getMetadata: () => extractMetadata(workbook),
    toCSV: (sheet?: string | number) => worksheetToCSV(workbook, sheet),
    evaluateCell: (sheet, row, column) =>
      evaluateCellFormula(workbook, sheet, row, column),
    evaluateAll: () => evaluateWorkbook(workbook),
    findCircularReferences: () => detectCircularReferences(workbook),
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
    formula: cell.f,
    evaluatedValue: cell.v ?? null,
    style: cell.s ? mapFromSheetJSStyle(cell.s) : undefined,
  };
}

function setCellValue(
  workbook: XLSX.WorkBook,
  sheet: string | number,
  row: number,
  column: number,
  value: ExcelCellValue,
  options: ExcelSetCellOptions = {}
) {
  const sheetName = resolveSheetName(workbook, sheet);
  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) {
    throw new Error(`Sheet "${sheetName}" not found.`);
  }

  if (row < 1 || column < 1) {
    throw new Error("Row and column must be 1-based positive integers.");
  }

  const origin = { r: row - 1, c: column - 1 };

  const writeValue =
    options.formula !== undefined && options.formula !== ""
      ? `=${options.formula}`
      : value ?? null;

  XLSX.utils.sheet_add_aoa(
    worksheet,
    [[writeValue]],
    {
      origin,
    }
  );

  const cellAddress = XLSX.utils.encode_cell(origin);
  const cell = worksheet[cellAddress];
  if (cell) {
    if (options.formula) {
      cell.f = options.formula;
      if (writeValue === null) {
        delete cell.v;
      }
    } else if (cell.f && value !== undefined) {
      delete cell.f;
    }

    if (value instanceof Date) {
      cell.t = "d";
      cell.v = value;
    }

    if (options.style) {
      cell.s = {
        ...(cell.s ?? {}),
        ...mapToSheetJSStyle(options.style),
      };
    }
  }
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

type SheetJSStyle = XLSX.CellObject["s"];

function mapToSheetJSStyle(style: ExcelCellStyle): SheetJSStyle {
  const result: SheetJSStyle = {};

  if (
    style.bold !== undefined ||
    style.italic !== undefined ||
    style.underline !== undefined ||
    style.fontColor ||
    style.fontName ||
    style.fontSize
  ) {
    result.font = {
      bold: style.bold,
      italic: style.italic,
      underline: style.underline ? true : undefined,
      color: style.fontColor
        ? { rgb: normalizeColor(style.fontColor) }
        : undefined,
      name: style.fontName,
      sz: style.fontSize,
    };
  }

  if (style.backgroundColor) {
    result.fill = {
      patternType: "solid",
      fgColor: { rgb: normalizeColor(style.backgroundColor) },
    };
  }

  if (style.horizontalAlign || style.verticalAlign || style.numberFormat) {
    result.alignment = {
      horizontal: style.horizontalAlign,
      vertical: style.verticalAlign,
    };
  }

  if (style.numberFormat) {
    result.numFmt = style.numberFormat;
  }

  return result;
}

function mapFromSheetJSStyle(style: SheetJSStyle): ExcelCellStyle {
  const font = style?.font ?? {};
  const fill = style?.fill ?? {};
  const alignment = style?.alignment ?? {};

  return {
    bold: font.bold ?? undefined,
    italic: font.italic ?? undefined,
    underline: font.underline ? true : undefined,
    fontColor: font.color?.rgb,
    fontName: font.name,
    fontSize: typeof font.sz === "number" ? font.sz : undefined,
    backgroundColor: fill?.fgColor?.rgb,
    horizontalAlign: alignment.horizontal as
      | "left"
      | "center"
      | "right"
      | undefined,
    verticalAlign: alignment.vertical as
      | "top"
      | "center"
      | "bottom"
      | undefined,
    numberFormat: style?.numFmt,
  };
}

function normalizeColor(color: string): string {
  if (!color) {
    return "000000";
  }

  let hex = color.trim();
  if (hex.startsWith("#")) {
    hex = hex.slice(1);
  }

  if (hex.length === 3) {
    hex = hex
      .split("")
      .map((char) => char + char)
      .join("");
  }

  return hex.padEnd(6, "0").slice(0, 6).toUpperCase();
}

function evaluateWorkbook(workbook: XLSX.WorkBook) {
  try {
    XLSX_CALC(workbook);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new Error(`Formula evaluation failed: ${message}`);
  }
}

function evaluateCellFormula(
  workbook: XLSX.WorkBook,
  sheet: string | number,
  row: number,
  column: number
) {
  evaluateWorkbook(workbook);
  const cell = getCellValue(workbook, sheet, row, column);
  return {
    address: cell.address,
    formula: cell.formula,
    value: cell.value,
    type: cell.type,
  };
}

function detectCircularReferences(workbook: XLSX.WorkBook) {
  const graph = buildFormulaGraph(workbook);
  const visited = new Set<string>();
  const stack = new Set<string>();
  const path: string[] = [];
  const cycles: ExcelCircularReference[] = [];

  const dfs = (node: string) => {
    if (stack.has(node)) {
      const cycleStart = path.indexOf(node);
      if (cycleStart !== -1) {
        const cyclePath = path.slice(cycleStart).concat(node);
        cycles.push({ path: cyclePath });
      }
      return;
    }

    if (visited.has(node)) {
      return;
    }

    visited.add(node);
    stack.add(node);
    path.push(node);

    const dependencies = graph.get(node);
    if (dependencies) {
      dependencies.forEach((next) => dfs(next));
    }

    path.pop();
    stack.delete(node);
  };

  graph.forEach((_deps, node) => {
    if (!visited.has(node)) {
      dfs(node);
    }
  });

  return cycles;
}

function buildFormulaGraph(workbook: XLSX.WorkBook) {
  const graph = new Map<string, Set<string>>();

  for (const sheetName of workbook.SheetNames) {
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) continue;

    for (const [address, cell] of Object.entries(worksheet)) {
      if (!cell || address.startsWith("!")) continue;
      if (!cell.f) continue;

      const node = `${sheetName}!${normalizeAddress(address)}`;
      const refs = extractCellReferences(cell.f, sheetName);
      if (!graph.has(node)) {
        graph.set(node, new Set());
      }
      const deps = graph.get(node)!;
      refs.forEach((ref) => deps.add(ref));
    }
  }

  return graph;
}

function extractCellReferences(formula: string, currentSheet: string) {
  const matches = formula.matchAll(
    /(?:'([^']+)'|([A-Za-z0-9_]+))?!?\$?([A-Z]{1,3})\$?([0-9]+)/g
  );

  const references: string[] = [];
  for (const match of matches) {
    const sheetRef = match[1] ?? match[2] ?? currentSheet;
    const column = match[3];
    const row = match[4];
    references.push(`${sheetRef}!${column}${row}`);
  }

  return references.map(normalizeReference);
}

function normalizeReference(reference: string) {
  const [sheet, address] = reference.split("!");
  return `${sheet}!${normalizeAddress(address)}`;
}

function normalizeAddress(address: string) {
  return address.replace(/\$/g, "").toUpperCase();
}

