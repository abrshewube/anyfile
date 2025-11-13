import type { FileMetadata } from "@anyfile/core";
import type { WorkBook } from "xlsx";

export interface ExcelWorksheetDescriptor {
  name: string;
  range?: string;
  rowCount: number;
  columnCount: number;
}

export interface ExcelReadOptions {
  sheet?: string | number;
  headerRow?: number;
  range?: string;
}

export interface ExcelWriteOptions {
  sheetName?: string;
}

export interface ExcelFileData {
  workbook: WorkBook;
  worksheets: ExcelWorksheetDescriptor[];
  readSheet: (
    nameOrIndex?: string | number,
    options?: ExcelReadOptions
  ) => Promise<Record<string, unknown>[]>;
  toJSON: () => Promise<Record<string, Record<string, unknown>[]>>;
}

export interface ExcelOpenOptions {
  readOptions?: ExcelReadOptions;
  metadata?: Partial<FileMetadata>;
}

