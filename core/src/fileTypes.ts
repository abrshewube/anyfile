export type FileType =
  | "excel"
  | "pdf"
  | "word"
  | "csv"
  | "ppt"
  | "text"
  | "image"
  | "archive";

export interface FileMetadata {
  name: string;
  size: number;
  type: FileType;
  createdAt?: Date;
  modifiedAt?: Date;
  mimeType?: string;
}

