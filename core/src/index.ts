import type { FileMetadata, FileType } from "./fileTypes";
import type { AnyFileSource, FileHandler } from "./registry";
import {
  getHandlerByExtension,
  getHandlerByType,
  listRegisteredHandlers,
  registerFileType,
} from "./registry";
import { getExtensionFromPath, getFileName } from "./utils";

export type { FileMetadata, FileType } from "./fileTypes";
export type {
  AnyFileSource,
  FileHandler,
  FileHandlerContext,
  FileHandlerResult,
} from "./registry";
export interface OpenOptions {
  type?: FileType;
  metadata?: Partial<FileMetadata>;
}

export interface AnyFileInstance<TData = unknown> {
  type: FileType;
  metadata: FileMetadata;
  read: () => Promise<TData>;
  write: (outputPath: string, data: TData) => Promise<void>;
  convert?: <TNext = unknown>(toType: FileType) => Promise<AnyFileInstance<TNext>>;
}

export const AnyFile = {
  register<TData>(handler: FileHandler<TData>) {
    registerFileType(handler);
    return handler;
  },

  async open<TData = unknown>(
    source: AnyFileSource,
    options: OpenOptions = {}
  ): Promise<AnyFileInstance<TData>> {
    const handler = await resolveHandler(source, options);
    if (!handler) {
      throw new Error("No handler registered for the provided source.");
    }

    const result = await handler.open({
      source,
      metadata: buildInitialMetadata(source, handler, options.metadata),
    });

    return {
      type: result.type,
      metadata: result.metadata,
      read: result.read as () => Promise<TData>,
      write: result.write as (outputPath: string, data: TData) => Promise<void>,
      convert: result.convert
        ? async <TNext = unknown>(toType: FileType) => {
            const converted = await result.convert!(toType);
            return {
              type: converted.type,
              metadata: converted.metadata,
              read: converted.read,
              write: converted.write,
              convert: converted.convert,
            } as AnyFileInstance<TNext>;
          }
        : undefined,
    } as AnyFileInstance<TData>;
  },

  getHandler(type: FileType): FileHandler | undefined {
    return getHandlerByType(type);
  },
};

async function resolveHandler(
  source: AnyFileSource,
  options: OpenOptions
): Promise<FileHandler | undefined> {
  if (options.type) {
    const handler = getHandlerByType(options.type);
    if (!handler) {
      throw new Error(`No handler registered for type "${options.type}".`);
    }
    return handler;
  }

  if (typeof source === "string") {
    const extension = getExtensionFromPath(source);
    if (!extension) {
      throw new Error(
        "Unable to determine file type from path. Provide a type explicitly in open options."
      );
    }

    const handler = getHandlerByExtension(extension);
    if (handler) {
      return handler;
    }

    throw new Error(
      `No handler registered for files with extension ".${extension}".`
    );
  }

  const detected = await detectFromBuffer(source);
  if (detected.length === 1) {
    return detected[0];
  }

  if (detected.length > 1) {
    throw new Error(
      "Multiple handlers match the provided source. Provide a type explicitly in open options."
    );
  }

  return undefined;
}

async function detectFromBuffer(
  source: Exclude<AnyFileSource, string>
): Promise<FileHandler[]> {
  const matches: FileHandler[] = [];
  for (const handler of listRegisteredHandlers()) {
    if (!handler.detect) {
      continue;
    }

    const result = await handler.detect(source);
    if (result) {
      matches.push(handler);
    }
  }

  return matches;
}

function buildInitialMetadata(
  source: AnyFileSource,
  handler: FileHandler,
  metadata: Partial<FileMetadata> = {}
): Partial<FileMetadata> {
  const base: Partial<FileMetadata> = {
    type: handler.type,
    ...metadata,
  };

  if (typeof source === "string") {
    base.name = metadata.name ?? getFileName(source);
  }

  return base;
}

