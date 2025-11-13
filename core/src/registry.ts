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
  write: (output: string, data: TData) => Promise<void>;
  convert?: <TNext = unknown>(to: FileType) => Promise<FileHandlerResult<TNext>>;
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

export function registerFileType<TData>(handler: FileHandler<TData>) {
  if (registry.byType.has(handler.type)) {
    throw new Error(`Handler for type "${handler.type}" is already registered.`);
  }

  registry.byType.set(handler.type, handler as FileHandler);
  handler.extensions.forEach((ext) => {
    if (registry.byExtension.has(ext.toLowerCase())) {
      throw new Error(
        `Handler for extension ".${ext.toLowerCase()}" is already registered.`
      );
    }
    registry.byExtension.set(ext.toLowerCase(), handler as FileHandler);
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

export function clearRegistry() {
  registry.byType.clear();
  registry.byExtension.clear();
}

