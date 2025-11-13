import { beforeEach, describe, expect, it } from "vitest";

import { AnyFile } from "./index";
import type { FileHandler } from "./registry";
import { clearRegistry } from "./registry";

const createTestHandler = (): FileHandler<string> => ({
  type: "text",
  extensions: ["txt", "md"],
  detect(source) {
    if (typeof source === "string") {
      return source.endsWith(".txt") || source.endsWith(".md");
    }

    if (source instanceof Uint8Array || source instanceof ArrayBuffer) {
      return true;
    }

    return false;
  },
  async open({ source, metadata }) {
    const finalMetadata = {
      name:
        metadata?.name ??
        (typeof source === "string" ? source.split(/[\\/]/).pop() ?? "unknown.txt" : "buffer.txt"),
      size: metadata?.size ?? 4,
      type: "text" as const,
      createdAt: metadata?.createdAt,
      modifiedAt: metadata?.modifiedAt,
    };

    return {
      type: "text",
      metadata: finalMetadata,
      read: async () => "test",
      write: async (_output, _data) => {},
      convert: async <TNext = string>(toType: "pdf" | "text") => ({
        type: toType,
        metadata: { ...finalMetadata, type: toType },
        read: async () => `converted-${toType}`,
        write: async (_output, _data) => {},
      }),
    };
  },
});

describe("AnyFile core", () => {
  beforeEach(() => {
    clearRegistry();
    AnyFile.register(createTestHandler());
  });

  it("opens a file by inferring handler from extension", async () => {
    const file = await AnyFile.open("example.txt");
    expect(file.type).toBe("text");
    await expect(file.read()).resolves.toBe("test");
  });

  it("allows conversion using the handler-provided convert function", async () => {
    const file = await AnyFile.open("example.txt");
    const converted = await file.convert?.("pdf");
    expect(converted?.type).toBe("pdf");
    await expect(converted?.read()).resolves.toBe("converted-pdf");
  });

  it("throws when no handler is registered for the extension", async () => {
    await expect(AnyFile.open("unknown.xyz")).rejects.toThrow(
      'No handler registered for files with extension ".xyz".'
    );
  });

  it("uses provided metadata overrides when opening", async () => {
    const file = await AnyFile.open("example.txt", {
      metadata: { name: "Custom.txt", size: 10 },
    });
    expect(file.metadata.name).toBe("Custom.txt");
    expect(file.metadata.size).toBe(10);
  });
});

