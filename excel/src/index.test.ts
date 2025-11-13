import { beforeAll, describe, expect, it } from "vitest";

import { AnyFile } from "@anyfile/core";
import * as XLSX from "xlsx";

import { Excel } from "./index";
import type { ExcelFileData } from "./types";

const SAMPLE_ROWS = [
  { Name: "Alice", Score: 95 },
  { Name: "Bob", Score: 87 },
];

function createWorkbookBuffer() {
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(SAMPLE_ROWS);
  XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
  return XLSX.write(workbook, { bookType: "xlsx", type: "buffer" }) as Buffer;
}

describe("@anyfile/excel handler", () => {
  beforeAll(() => {
    Excel.register();
  });

  it("detects excel sources by extension and signature", async () => {
    expect(await Excel.detect("report.xlsx")).toBe(true);
    expect(await Excel.detect("notes.txt")).toBe(false);

    const buffer = createWorkbookBuffer();
    expect(await Excel.detect(buffer)).toBe(true);
  });

  it("opens a workbook from a buffer and reads rows", async () => {
    const buffer = createWorkbookBuffer();
    const file = await Excel.open(buffer);

    expect(file.metadata.type).toBe("excel");
    expect(file.worksheets[0]?.name).toBe("Sheet1");

    const rows = await file.readSheet("Sheet1");
    expect(rows).toEqual(SAMPLE_ROWS);
  });

  it("integrates with AnyFile.open", async () => {
    const buffer = createWorkbookBuffer();
    const anyFile = await AnyFile.open<ExcelFileData>(buffer, { type: "excel" });
    const data = await anyFile.read();

    expect(data.worksheets.length).toBeGreaterThan(0);
    expect(await data.readSheet("Sheet1")).toEqual(SAMPLE_ROWS);
  });
});

