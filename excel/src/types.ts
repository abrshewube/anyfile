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
    value: ExcelCellValue,
    options?: ExcelSetCellOptions
  ) => void;
  addSheet: (name: string) => void;
  addSheetFromCSV: (name: string, csv: string) => void;
  deleteSheet: (name: string) => void;
  getMetadata: () => ExcelMetadata;
  toCSV: (sheet?: string | number) => string;
  evaluateCell: (
    sheet: string | number,
    row: number,
    column: number
  ) => ExcelFormulaResult;
  evaluateAll: () => void;
  findCircularReferences: () => ExcelCircularReference[];
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
  formula?: string;
  evaluatedValue?: ExcelCellValue;
  style?: ExcelCellStyle;
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

export interface ExcelSetCellOptions {
  formula?: string;
  style?: ExcelCellStyle;
}

export interface ExcelCellStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontColor?: string;
  fontName?: string;
  fontSize?: number;
  backgroundColor?: string;
  horizontalAlign?: "left" | "center" | "right";
  verticalAlign?: "top" | "center" | "bottom";
  numberFormat?: string;
}

export interface ExcelFormulaResult {
  address: string;
  formula?: string;
  value: ExcelCellValue;
  type?: string;
}

export interface ExcelCircularReference {
  path: string[];
}

