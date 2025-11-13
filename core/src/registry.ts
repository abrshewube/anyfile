import type { FileMetadata, FileType } from "./fileTypes";

export type AnyFileSource = string | ArrayBuffer | ArrayBufferView;

export interface FileHandlerContext {
  source: AnyFileSource;
  metadata?: Partial<FileMetadata>;
}

export interface FileHandlerResult<TData = unknown> {
  type: FileType;
  metadata: FileMetadata;
  read: () => Promise<TData>;
  write: (output: string, data?: TData) => Promise<void>;
  convert?: (to: FileType) => Promise<FileHandlerResult>;
}

export interface FileHandler<TData = unknown> {
  type: FileType;
  extensions: string[];
  detect?: (source: AnyFileSource) => Promise<boolean> | boolean;
  open: (context: FileHandlerContext) => Promise<FileHandlerResult<TData>>;
}

type HandlerRegistry = {
  byType: Map<FileType, FileHandler>;
  byExtension: Map<string, FileHandler>;
};

const registry: HandlerRegistry = {
  byType: new Map(),
  byExtension: new Map(),
};

export function registerFileType(handler: FileHandler) {
  registry.byType.set(handler.type, handler);
  handler.extensions.forEach((ext) => {
    registry.byExtension.set(ext.toLowerCase(), handler);
  });
}

export function getHandlerByType(type: FileType): FileHandler | undefined {
  return registry.byType.get(type);
}

export function getHandlerByExtension(extension: string): FileHandler | undefined {
  return registry.byExtension.get(extension.toLowerCase());
}

export function listRegisteredHandlers(): FileHandler[] {
  return [...registry.byType.values()];
}

