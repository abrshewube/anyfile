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
  convert?: <TNext = unknown>(
    to: FileType,
    options?: ConversionOptions
  ) => Promise<FileHandlerResult<TNext>>;
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

type ConversionKey = `${FileType}->${FileType}`;

export type ConversionOptions = Record<string, unknown> | undefined;

export type FileConversionHandler = <TInput = unknown, TOutput = unknown>(
  result: FileHandlerResult<TInput>,
  options?: ConversionOptions
) => Promise<FileHandlerResult<TOutput>>;

const conversionRegistry = new Map<ConversionKey, FileConversionHandler>();

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
  conversionRegistry.clear();
}

function toConversionKey(from: FileType, to: FileType): ConversionKey {
  return `${from}->${to}`;
}

export function registerConversion(
  from: FileType,
  to: FileType,
  handler: FileConversionHandler
) {
  const key = toConversionKey(from, to);
  conversionRegistry.set(key, handler);
}

export function getConversion(
  from: FileType,
  to: FileType
): FileConversionHandler | undefined {
  return conversionRegistry.get(toConversionKey(from, to));
}

export function listConversions(): Array<{ from: FileType; to: FileType }> {
  return Array.from(conversionRegistry.keys()).map((key) => {
    const [from, to] = key.split("->") as [FileType, FileType];
    return { from, to };
  });
}

