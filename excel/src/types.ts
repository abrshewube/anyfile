import type { FileMetadata } from "@anyfile/core";
import type { WorkBook, WorkSheet } from "xlsx";

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
  getSheets: () => ExcelWorksheetDescriptor[];
  getSheetNames: () => string[];
  readSheet: (
    nameOrIndex?: string | number,
    options?: ExcelReadOptions
  ) => Promise<Record<string, unknown>[]>;
  getCell: (
    sheet: string | number,
    row: number,
    column: number
  ) => ExcelCell | undefined;
  setCell: (
    sheet: string | number,
    row: number,
    column: number,
    value: ExcelCellValue
  ) => void;
  addSheet: (name: string) => void;
  addSheetFromCSV: (name: string, csv: string) => void;
  deleteSheet: (name: string) => void;
  getMetadata: () => ExcelMetadata;
  toCSV: (sheet?: string | number) => string;
  worksheets: ExcelWorksheetDescriptor[];
  toJSON: () => Promise<Record<string, Record<string, unknown>[]>>;
}

export interface ExcelOpenOptions {
  readOptions?: ExcelReadOptions;
  metadata?: Partial<FileMetadata>;
}

export type ExcelCellValue = string | number | boolean | Date | null | undefined;

export interface ExcelCell {
  address: string;
  value: ExcelCellValue;
  raw?: unknown;
  type?: string;
}

export interface ExcelMetadata {
  title?: string;
  subject?: string;
  author?: string;
  manager?: string;
  company?: string;
  category?: string;
  keywords?: string;
  comments?: string;
  lastAuthor?: string;
  sheetCount: number;
  createdAt?: Date;
  modifiedAt?: Date;
}

export type WorksheetResolver = (
  workbook: WorkBook,
  sheet: string | number
) => WorkSheet;

