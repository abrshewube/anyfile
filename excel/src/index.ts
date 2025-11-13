import { promises as fs } from "node:fs";
import {
  AnyFile,
  type AnyFileInstance,
  type AnyFileSource,
  type FileMetadata,
} from "@anyfile/core";

import {
  configureFormulaLocalization,
  createExcelHandler,
  registerCustomFormula,
  registerCustomFormulas,
} from "./handler";
import type {
  ExcelCsvConversionOptions,
  ExcelFileData,
  ExcelFormulaImplementation,
  ExcelFormulaMap,
  ExcelOpenOptions,
  ExcelPdfConversionOptions,
  ExcelReadOptions,
} from "./types";

const handler = createExcelHandler();
let registered = false;
let conversionsRegistered = false;

function ensureRegistered() {
  if (registered) {
    return handler;
  }

  try {
    AnyFile.register(handler);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (!message.includes("already registered")) {
      throw error;
    }
  }

  registered = true;
  registerExcelConversions();
  return handler;
}

function replaceExtension(name: string | undefined, next: string) {
  if (!name) {
    return `converted.${next}`;
  }
  return name.replace(/\.[^.]+$/, "") + `.${next}`;
}

function registerExcelConversions() {
  if (conversionsRegistered) {
    return;
  }
  conversionsRegistered = true;

  AnyFile.registerConversion("excel", "csv", async (result, options) => {
    const csvOptions = (options ?? {}) as ExcelCsvConversionOptions;
    const data = await result.read();
    const targetSheet =
      csvOptions.sheet ?? data.getSheetNames()[0] ?? "Sheet1";
    const delimiter = csvOptions.delimiter ?? ",";
    const csv = data.toCSV(targetSheet);
    const normalizedCsv =
      delimiter === ","
        ? csv
        : csv
            .split("\n")
            .map((line) => line.replace(/,/g, delimiter))
            .join("\n");

    const metadata: FileMetadata = {
      ...result.metadata,
      type: "csv",
      name: replaceExtension(result.metadata.name, "csv"),
      mimeType: "text/csv",
      size: Buffer.byteLength(normalizedCsv, "utf8"),
    };

    return {
      type: "csv",
      metadata,
      read: async () => normalizedCsv,
      write: async (output, content) => {
        const toWrite = content ?? normalizedCsv;
        await fs.writeFile(output, toWrite, "utf8");
      },
    };
  });

  AnyFile.registerConversion("excel", "pdf", async (result, options) => {
    const pdfOptions = (options ?? {}) as ExcelPdfConversionOptions;
    if (!pdfOptions.renderer) {
      throw new Error(
        "Excel PDF conversion requires a renderer option (renderer: (workbook) => Promise<Uint8Array>)."
      );
    }

    const data = await result.read();
    const buffer = await pdfOptions.renderer(data.workbook);

    const metadata: FileMetadata = {
      ...result.metadata,
      type: "pdf",
      name: replaceExtension(result.metadata.name, "pdf"),
      mimeType: "application/pdf",
      size: buffer.byteLength,
    };

    return {
      type: "pdf",
      metadata,
      read: async () => buffer,
      write: async (output, content) => {
        const toWrite = content ?? buffer;
        await fs.writeFile(output, Buffer.from(toWrite));
      },
    };
  });
}

export interface ExcelFileHandle extends ExcelFileData {
  metadata: FileMetadata;
  write: (outputPath: string, data?: ExcelFileData) => Promise<void>;
  convert?: AnyFileInstance<ExcelFileData>["convert"];
}

export const Excel = {
  register: () => ensureRegistered(),

  registerFormula: (
    name: string,
    implementation: ExcelFormulaImplementation
  ) => {
    ensureRegistered();
    registerCustomFormula(name, implementation);
  },

  registerFormulas: (formulas: ExcelFormulaMap) => {
    ensureRegistered();
    registerCustomFormulas(formulas);
  },

  configureLocalization: (localization: Record<string, string>) => {
    ensureRegistered();
    configureFormulaLocalization(localization);
  },

  detect: async (source: AnyFileSource) => {
    return ensureRegistered().detect?.(source) ?? false;
  },

  async open(
    source: AnyFileSource,
    options: ExcelOpenOptions = {}
  ): Promise<ExcelFileHandle> {
    ensureRegistered();
    const file = await AnyFile.open<ExcelFileData>(source, {
      type: "excel",
      metadata: options.metadata,
    });

    const data = await file.read();
    const readSheet = (
      nameOrIndex?: string | number,
      readOptions?: ExcelReadOptions
    ) =>
      data.readSheet(nameOrIndex, {
        ...options.readOptions,
        ...readOptions,
      });

    return {
      ...data,
      readSheet,
      metadata: file.metadata,
      write: async (outputPath, payload = data) => file.write(outputPath, payload),
      convert: file.convert,
    };
  },
};

ensureRegistered();

export type {
  ExcelFileData,
  ExcelOpenOptions,
  ExcelReadOptions,
  ExcelWorksheetDescriptor,
  ExcelWriteOptions,
  ExcelCell,
  ExcelCellValue,
  ExcelCellStyle,
  ExcelCircularReference,
  ExcelEvaluationReport,
  ExcelEvaluateAllOptions,
  ExcelFormulaResult,
  ExcelFormulaSummary,
  ExcelChartSummary,
  ExcelImageSummary,
  ExcelMacroSummary,
  ExcelMetadata,
  ExcelSetCellOptions,
  ExcelFormulaImplementation,
  ExcelFormulaMap,
  ExcelCsvConversionOptions,
  ExcelPdfConversionOptions,
} from "./types";

