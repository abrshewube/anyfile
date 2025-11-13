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
  ExcelFileData,
  ExcelFormulaImplementation,
  ExcelFormulaMap,
  ExcelOpenOptions,
  ExcelReadOptions,
} from "./types";

const handler = createExcelHandler();
let registered = false;

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
  return handler;
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
  ExcelMetadata,
  ExcelSetCellOptions,
  ExcelFormulaImplementation,
  ExcelFormulaMap,
} from "./types";

