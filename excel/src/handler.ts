import { promises as fs } from "node:fs";
import { basename, dirname, extname } from "node:path";

import type {
  AnyFileSource,
  FileHandler,
  FileHandlerContext,
  FileHandlerResult,
  FileMetadata,
  FileType,
} from "@anyfile/core";
import JSZip from "jszip";
import * as XLSX from "xlsx";
import XLSX_CALC from "xlsx-calc";

import type {
  ExcelCellStyle,
  ExcelCellValue,
  ExcelCircularReference,
  ExcelChartSummary,
  ExcelChartSeries,
  ExcelEvaluateAllOptions,
  ExcelEvaluationReport,
  ExcelFileData,
  ExcelFormulaImplementation,
  ExcelFormulaMap,
  ExcelFormulaResult,
  ExcelFormulaSummary,
  ExcelImageSummary,
  ExcelMacroSummary,
  ExcelMetadata,
  ExcelReadOptions,
  ExcelSetCellOptions,
  ExcelWorksheetDescriptor,
} from "./types";

const EXCEL_EXTENSIONS = ["xls", "xlsx", "xlsm", "xlsb"];
const XLS_SIGNATURE = Buffer.from([0xd0, 0xcf, 0x11, 0xe0]);
const XLSX_SIGNATURE = Buffer.from([0x50, 0x4b, 0x03, 0x04]);
const customFormulaRegistry = new Set<string>();
let lastEvaluationTimestamp: Date | undefined;

export function createExcelHandler(): FileHandler<ExcelFileData> {
  return {
    type: "excel",
    extensions: EXCEL_EXTENSIONS,
    detect: async (source: AnyFileSource) => detectExcelSource(source),
    open: async ({ source, metadata }: FileHandlerContext) => {
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

      const assets = createAssetExtractors(payload.buffer, workbook);
      return buildHandlerResult(workbook, fileMetadata, assets);
    },
  };
}

type AssetExtractors = {
  getCharts: () => Promise<ExcelChartSummary[]>;
  getImages: () => Promise<ExcelImageSummary[]>;
  getMacros: () => Promise<ExcelMacroSummary[]>;
};

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
  metadata: FileMetadata,
  assets: AssetExtractors
): FileHandlerResult<ExcelFileData> {
  const createData = () => createExcelFileData(workbook, assets);

  return {
    type: "excel",
    metadata,
    read: async () => createData(),
    write: async (outputPath: string, data: ExcelFileData) => {
      const workbookToWrite = data.workbook ?? workbook;
      const bookType = resolveBookType(outputPath);
      const buffer = XLSX.write(workbookToWrite, {
        bookType,
        type: "buffer",
      });
      await fs.writeFile(outputPath, buffer);
    },
    convert: async <TNext = unknown>(
      toType: FileType,
      conversionOptions?: unknown
    ): Promise<FileHandlerResult<TNext>> => {
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
        read: async () => csv as unknown as TNext,
        write: async (output: string, data: TNext) => {
          const content = (data as unknown as string) ?? csv;
          await fs.writeFile(output, content, "utf8");
        },
      };
    },
  };
}

function createExcelFileData(
  workbook: XLSX.WorkBook,
  assets: AssetExtractors
): ExcelFileData {
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

  const createReadSheetStream = (
    nameOrIndex: string | number = 0,
    streamOptions?: ExcelReadOptions & { chunkSize?: number }
  ) =>
    (async function* () {
      const rows = await readSheet(nameOrIndex, streamOptions);
      const chunkSize = streamOptions?.chunkSize ?? 256;

      for (let index = 0; index < rows.length; index += 1) {
        yield rows[index];

        if ((index + 1) % chunkSize === 0) {
          await Promise.resolve();
        }
      }
    })();

  const api: ExcelFileData = {
    workbook,
    getSheets: describeSheets,
    getSheetNames: () => workbook.SheetNames.slice(),
    readSheet,
    readSheetStream: (nameOrIndex, streamOptions) =>
      createReadSheetStream(nameOrIndex, streamOptions),
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
    evaluateAll: (options) => evaluateWorkbook(workbook, options),
    findCircularReferences: () => detectCircularReferences(workbook),
    getFormulaSummary: () => summarizeFormulas(workbook),
    getCharts: () => assets.getCharts(),
    getImages: () => assets.getImages(),
    listMacros: () => assets.getMacros(),
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

function evaluateWorkbook(
  workbook: XLSX.WorkBook,
  options: ExcelEvaluateAllOptions = {}
): ExcelEvaluationReport {
  const report: ExcelEvaluationReport = {
    evaluated: [],
    circular: [],
    errors: [],
  };

  let graph: Map<string, Set<string>> | undefined;

  try {
    XLSX_CALC(workbook);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (message.toLowerCase().includes("circular")) {
      graph = buildFormulaGraph(workbook);
      report.circular = detectCircularReferences(workbook, graph);
      if (!options.ignoreCircular) {
        throw new Error(`Formula evaluation failed: ${message}`);
      }
    } else {
      report.errors.push({ address: "", message });
      throw new Error(`Formula evaluation failed: ${message}`);
    }
  }

  if (!graph) {
    graph = buildFormulaGraph(workbook);
  }

  report.evaluated = [...graph.keys()];

  if (report.circular.length === 0) {
    report.circular = detectCircularReferences(workbook, graph);
  }

  lastEvaluationTimestamp = new Date();

  return report;
}

function evaluateCellFormula(
  workbook: XLSX.WorkBook,
  sheet: string | number,
  row: number,
  column: number
): ExcelFormulaResult {
  const sheetName = resolveSheetName(workbook, sheet);
  const address = `${sheetName}!${XLSX.utils.encode_cell({
    r: row - 1,
    c: column - 1,
  })}`;

  const baseCell = getCellValue(workbook, sheet, row, column);

  try {
    const report = evaluateWorkbook(workbook, { ignoreCircular: true });
    const evaluatedCell = getCellValue(workbook, sheet, row, column);

    if (!evaluatedCell) {
      return {
        address,
        sheet: sheetName,
        formula: baseCell?.formula,
        value: null,
        evaluatedValue: null,
        type: undefined,
        error: "Cell not found.",
      };
    }

    const isCircular = report.circular.some((entry) =>
      entry.path.includes(address)
    );

    return {
      address,
      sheet: sheetName,
      formula: evaluatedCell.formula ?? baseCell?.formula,
      value: evaluatedCell.value,
      evaluatedValue: evaluatedCell.evaluatedValue ?? evaluatedCell.value,
      type: evaluatedCell.type,
      error: isCircular ? "Circular reference detected." : undefined,
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return {
      address,
      sheet: sheetName,
      formula: baseCell?.formula,
      value: baseCell?.value ?? null,
      evaluatedValue: baseCell?.evaluatedValue ?? null,
      type: baseCell?.type,
      error: message,
    };
  }
}

function detectCircularReferences(
  workbook: XLSX.WorkBook,
  precomputedGraph?: Map<string, Set<string>>
) {
  const graph = precomputedGraph ?? buildFormulaGraph(workbook);
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

function summarizeFormulas(workbook: XLSX.WorkBook): ExcelFormulaSummary {
  const graph = buildFormulaGraph(workbook);
  const circular = detectCircularReferences(workbook, graph);

  const sheets = new Set<string>();
  graph.forEach((_deps, node) => {
    const [sheet] = node.split("!");
    sheets.add(sheet);
  });

  return {
    totalFormulas: graph.size,
    sheetsWithFormulas: sheets.size,
    circularReferences: circular.length,
    lastEvaluatedAt: lastEvaluationTimestamp,
    customFormulas: Array.from(customFormulaRegistry).sort(),
  };
}

function createAssetExtractors(
  buffer: Buffer,
  workbook: XLSX.WorkBook
): AssetExtractors {
  let zipPromise: Promise<JSZip | null> | undefined;
  const getZip = async () => {
    if (!zipPromise) {
      zipPromise = JSZip.loadAsync(buffer).catch(() => null);
    }
    return zipPromise;
  };

  let sheetEntriesPromise:
    | Promise<Array<{ name: string; path: string }>>
    | undefined;

  const getSheetEntries = async () => {
    if (!sheetEntriesPromise) {
      sheetEntriesPromise = (async () => {
        const zip = await getZip();
        if (!zip) {
          return workbook.SheetNames.map((name, index) => ({
            name,
            path: `worksheets/sheet${index + 1}.xml`,
          }));
        }

        const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
        const relsXml = await zip
          .file("xl/_rels/workbook.xml.rels")
          ?.async("string");
        const rels = parseRelationships(relsXml);
        const entries: Array<{ name: string; path: string }> = [];

        if (workbookXml) {
          const sheetRegex = /<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/g;
          let match: RegExpExecArray | null;
          while ((match = sheetRegex.exec(workbookXml)) !== null) {
            const name = match[1];
            let target = rels.get(match[2]) ?? "";
            if (target.startsWith("/")) {
              target = target.slice(1);
            }
            if (!target) {
              const fallbackIndex = entries.length + 1;
              target = `worksheets/sheet${fallbackIndex}.xml`;
            }
            entries.push({ name, path: target });
          }
        }

        if (!entries.length) {
          return workbook.SheetNames.map((name, index) => ({
            name,
            path: `worksheets/sheet${index + 1}.xml`,
          }));
        }

        return entries;
      })();
    }

    return sheetEntriesPromise;
  };

  let chartsPromise: Promise<ExcelChartSummary[]> | undefined;
  let imagesPromise: Promise<ExcelImageSummary[]> | undefined;
  let macrosPromise: Promise<ExcelMacroSummary[]> | undefined;

  return {
    getCharts: async () => (chartsPromise ??= loadCharts()),
    getImages: async () => (imagesPromise ??= loadImages()),
    getMacros: async () => (macrosPromise ??= loadMacros()),
  };

  async function loadCharts(): Promise<ExcelChartSummary[]> {
    const zip = await getZip();
    if (!zip) {
      return [];
    }

    const sheets = await getSheetEntries();
    const chartTypeCache = new Map<string, string | undefined>();
    const results: ExcelChartSummary[] = [];

    for (const sheet of sheets) {
      const sheetXmlFile = zip.file(normalizeZipPath(sheet.path));
      if (!sheetXmlFile) {
        continue;
      }

      const sheetXml = await sheetXmlFile.async("string");
      const drawingMatches = [
        ...sheetXml.matchAll(/<drawing[^>]*r:id="([^"]+)"/g),
      ];
      if (!drawingMatches.length) {
        continue;
      }

      const sheetRelPath = buildSheetRelPath(sheet.path);
      const sheetRelXml = await zip.file(normalizeZipPath(sheetRelPath))?.async("string");
      const sheetRels = parseRelationships(sheetRelXml);

      for (const match of drawingMatches) {
        const drawingRelId = match[1];
        const drawingTarget = sheetRels.get(drawingRelId);
        if (!drawingTarget) {
          continue;
        }

        const drawingPath = resolveTarget(sheet.path, drawingTarget);
        const drawingXmlFile = zip.file(normalizeZipPath(drawingPath));
        if (!drawingXmlFile) {
          continue;
        }

        const drawingXml = await drawingXmlFile.async("string");
        const drawingRelPath = buildDrawingRelPath(drawingPath);
        const drawingRelXml = await zip
          .file(normalizeZipPath(drawingRelPath))
          ?.async("string");
        const drawingRels = parseRelationships(drawingRelXml);
        const anchors = parseDrawingAnchors(drawingXml);

        for (const anchor of anchors) {
          if (!anchor.chartId) {
            continue;
          }
          const chartTarget = drawingRels.get(anchor.chartId);
          if (!chartTarget) {
            continue;
          }
          const chartPath = resolveTarget(drawingPath, chartTarget);
          const type = await loadChartType(zip, chartPath, chartTypeCache);
          results.push({
            sheet: sheet.name,
            name: anchor.name,
            type,
            cellRange: anchor.range,
          });
        }
      }
    }

    return results;
  }

  async function loadImages(): Promise<ExcelImageSummary[]> {
    const zip = await getZip();
    if (!zip) {
      return [];
    }

    const sheets = await getSheetEntries();
    const results: ExcelImageSummary[] = [];

    for (const sheet of sheets) {
      const sheetXmlFile = zip.file(normalizeZipPath(sheet.path));
      if (!sheetXmlFile) {
        continue;
      }

      const sheetXml = await sheetXmlFile.async("string");
      const drawingMatches = [
        ...sheetXml.matchAll(/<drawing[^>]*r:id="([^"]+)"/g),
      ];
      if (!drawingMatches.length) {
        continue;
      }

      const sheetRelPath = buildSheetRelPath(sheet.path);
      const sheetRelXml = await zip.file(normalizeZipPath(sheetRelPath))?.async("string");
      const sheetRels = parseRelationships(sheetRelXml);

      for (const match of drawingMatches) {
        const drawingRelId = match[1];
        const drawingTarget = sheetRels.get(drawingRelId);
        if (!drawingTarget) {
          continue;
        }

        const drawingPath = resolveTarget(sheet.path, drawingTarget);
        const drawingXmlFile = zip.file(normalizeZipPath(drawingPath));
        if (!drawingXmlFile) {
          continue;
        }

        const drawingXml = await drawingXmlFile.async("string");
        const drawingRelPath = buildDrawingRelPath(drawingPath);
        const drawingRelXml = await zip
          .file(normalizeZipPath(drawingRelPath))
          ?.async("string");
        const drawingRels = parseRelationships(drawingRelXml);
        const anchors = parseDrawingAnchors(drawingXml);

        for (const anchor of anchors) {
          if (!anchor.imageId) {
            continue;
          }
          const imageTarget = drawingRels.get(anchor.imageId);
          if (!imageTarget) {
            continue;
          }
          const imagePath = resolveTarget(drawingPath, imageTarget);
          results.push({
            sheet: sheet.name,
            name: anchor.name,
            position: anchor.range,
            mediaType: determineMediaType(imagePath),
          });
        }
      }
    }

    return results;
  }

  async function loadMacros(): Promise<ExcelMacroSummary[]> {
    const zip = await getZip();
    if (!zip) {
      return [];
    }

    const vbaFile = zip.file("xl/vbaProject.bin");
    if (!vbaFile) {
      return [];
    }

    try {
      const buffer = await vbaFile.async("nodebuffer");
      const text = buffer.toString("latin1");
      const modules = new Set<string>();

      const moduleRegex = /Module=([A-Za-z0-9_]+)/gi;
      let match: RegExpExecArray | null;
      while ((match = moduleRegex.exec(text)) !== null) {
        modules.add(match[1]);
      }

      if (text.includes("ThisWorkbook")) {
        modules.add("ThisWorkbook");
      }

      const sheetRegex = /Sheet\d+/gi;
      let sheetMatch: RegExpExecArray | null;
      while ((sheetMatch = sheetRegex.exec(text)) !== null) {
        modules.add(sheetMatch[0]);
      }

      if (!modules.size) {
        modules.add("VBAProject");
      }

      return Array.from(modules).map((name) => ({
        module: "VBAProject",
        name,
      }));
    } catch {
      return [
        {
          module: "VBAProject",
          name: "Unknown",
        },
      ];
    }
  }
}

export function registerCustomFormula(
  name: string,
  implementation: ExcelFormulaImplementation
) {
  if (!name || typeof implementation !== "function") {
    throw new Error("registerFormula requires a name and function implementation.");
  }

  XLSX_CALC.set_fx(name, implementation);
  customFormulaRegistry.add(name.toUpperCase());
}

export function registerCustomFormulas(formulas: ExcelFormulaMap) {
  if (!formulas) {
    return;
  }

  XLSX_CALC.import_functions(formulas);
  Object.keys(formulas).forEach((name) => {
    customFormulaRegistry.add(name.toUpperCase());
  });
}

export function configureFormulaLocalization(localization: Record<string, string>) {
  if (!localization) {
    return;
  }
  XLSX_CALC.localizeFunctions(localization);
}

function parseRelationships(xml?: string): Map<string, string> {
  const map = new Map<string, string>();
  if (!xml) {
    return map;
  }

  const regex = /<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g;
  let match: RegExpExecArray | null;
  while ((match = regex.exec(xml)) !== null) {
    map.set(match[1], match[2]);
  }
  return map;
}

function buildSheetRelPath(sheetPath: string): string {
  const dir = dirname(sheetPath);
  const file = sheetPath.substring(sheetPath.lastIndexOf("/") + 1);
  if (!dir) {
    return `_rels/${file}.rels`;
  }
  return `${dir}/_rels/${file}.rels`;
}

function buildDrawingRelPath(drawingPath: string): string {
  const dir = dirname(drawingPath);
  const file = drawingPath.substring(drawingPath.lastIndexOf("/") + 1);
  if (!dir) {
    return `_rels/${file}.rels`;
  }
  return `${dir}/_rels/${file}.rels`;
}

function resolveTarget(basePath: string, target: string): string {
  if (!target) {
    return target;
  }
  if (target.startsWith("/")) {
    return target.replace(/^\//, "");
  }

  const baseDir = dirname(basePath);
  const combined = (baseDir ? `${baseDir}/` : "") + target;
  const segments = combined.split("/");
  const stack: string[] = [];
  for (const segment of segments) {
    if (!segment || segment === ".") {
      continue;
    }
    if (segment === "..") {
      stack.pop();
    } else {
      stack.push(segment);
    }
  }
  return stack.join("/");
}

type DrawingAnchor = {
  name?: string;
  chartId?: string;
  imageId?: string;
  range?: string;
};

function parseDrawingAnchors(xml: string): DrawingAnchor[] {
  const anchors: DrawingAnchor[] = [];
  const anchorRegex = /<xdr:(?:twoCellAnchor|oneCellAnchor)[^>]*>([\s\S]*?)<\/xdr:(?:twoCellAnchor|oneCellAnchor)>/g;
  let match: RegExpExecArray | null;
  while ((match = anchorRegex.exec(xml)) !== null) {
    const block = match[0];
    const nameMatch = block.match(/<a:cNvPr[^>]*name="([^"]+)"/);
    const chartMatch = block.match(/<c:chart[^>]*r:id="([^"]+)"/);
    const imageMatch = block.match(/<a:blip[^>]*r:embed="([^"]+)"/);
    anchors.push({
      name: nameMatch?.[1],
      chartId: chartMatch?.[1],
      imageId: imageMatch?.[1],
      range: extractRangeFromAnchor(block),
    });
  }
  return anchors;
}

function extractRangeFromAnchor(block: string): string | undefined {
  const fromMatch = /<xdr:from>[\s\S]*?<xdr:col>(\d+)<\/xdr:col>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>[\s\S]*?<\/xdr:from>/.exec(
    block
  );
  if (!fromMatch) {
    return undefined;
  }
  const toMatch = /<xdr:to>[\s\S]*?<xdr:col>(\d+)<\/xdr:col>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>[\s\S]*?<\/xdr:to>/.exec(
    block
  );

  const fromCol = Number(fromMatch[1]);
  const fromRow = Number(fromMatch[2]);
  const toCol = toMatch ? Number(toMatch[1]) : fromCol;
  const toRow = toMatch ? Number(toMatch[2]) : fromRow;

  const start = `${columnNumberToName(fromCol)}${fromRow + 1}`;
  const end = `${columnNumberToName(toCol)}${toRow + 1}`;
  return toMatch ? `${start}:${end}` : start;
}

function columnNumberToName(index: number): string {
  let n = index + 1;
  let result = "";
  while (n > 0) {
    const remainder = (n - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

function normalizeZipPath(pathValue: string): string {
  let normalized = pathValue.replace(/\\/g, "/");
  normalized = normalized.replace(/^\.\//, "");
  if (normalized.startsWith("/")) {
    normalized = normalized.slice(1);
  }
  if (!normalized.startsWith("xl/")) {
    normalized = `xl/${normalized}`;
  }
  return normalized;
}

async function loadChartType(
  zip: JSZip,
  chartPath: string,
  cache: Map<string, string | undefined>
): Promise<string | undefined> {
  const normalizedPath = normalizeZipPath(chartPath);
  if (cache.has(normalizedPath)) {
    return cache.get(normalizedPath);
  }

  const chartFile = zip.file(normalizedPath);
  if (!chartFile) {
    cache.set(normalizedPath, undefined);
    return undefined;
  }

  const xml = await chartFile.async("string");
  const match = xml.match(/<c:([A-Za-z0-9]+)Chart\b/);
  const type = match ? match[1] : undefined;
  cache.set(normalizedPath, type);
  return type;
}

function determineMediaType(pathValue: string): string {
  const extension = extname(pathValue).toLowerCase();
  switch (extension) {
    case ".png":
      return "image/png";
    case ".jpg":
    case ".jpeg":
      return "image/jpeg";
    case ".gif":
      return "image/gif";
    case ".bmp":
      return "image/bmp";
    case ".svg":
      return "image/svg+xml";
    default:
      return "application/octet-stream";
  }
}

