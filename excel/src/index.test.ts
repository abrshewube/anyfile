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
    expect(file.getSheets()[0]?.name).toBe("Sheet1");

    const rows = await file.readSheet("Sheet1");
    expect(rows).toEqual(SAMPLE_ROWS);
  });

  it("integrates with AnyFile.open", async () => {
    const buffer = createWorkbookBuffer();
    const anyFile = await AnyFile.open<ExcelFileData>(buffer, { type: "excel" });
    const data = await anyFile.read();

    expect(data.getSheets().length).toBeGreaterThan(0);
    expect(await data.readSheet("Sheet1")).toEqual(SAMPLE_ROWS);
  });

  it("provides sheet and cell helpers", async () => {
    const buffer = createWorkbookBuffer();
    const file = await Excel.open(buffer);

    expect(file.getSheetNames()).toEqual(["Sheet1"]);
    expect(file.getMetadata().sheetCount).toBe(1);

    const cell = file.getCell("Sheet1", 1, 1);
    expect(cell?.value).toBe("Name");

    file.setCell("Sheet1", 4, 1, "Charlie");
    file.setCell("Sheet1", 4, 2, 72);

    const updatedRows = await file.readSheet("Sheet1", { headerRow: 1 });
    expect(updatedRows).toEqual([
      { Name: "Alice", Score: 95 },
      { Name: "Bob", Score: 87 },
      { Name: "Charlie", Score: 72 },
    ]);

    file.addSheet("Summary");
    expect(file.getSheetNames()).toEqual(["Sheet1", "Summary"]);

    file.deleteSheet("Summary");
    expect(file.getSheetNames()).toEqual(["Sheet1"]);
  });

  it("exports csv and converts via AnyFile", async () => {
    const buffer = createWorkbookBuffer();
    const file = await Excel.open(buffer);

    const csv = file.toCSV();
    expect(csv).toContain("Name,Score");
    expect(csv).toContain("Alice,95");

    file.addSheetFromCSV("Imported", "Col1,Col2\n1,2\n3,4");
    expect(file.getSheetNames()).toEqual(["Sheet1", "Imported"]);

    const anyFile = await AnyFile.open(buffer, { type: "excel" });
    const converted = await anyFile.convert?.("csv");
    expect(converted?.type).toBe("csv");
    expect(await converted?.read()).toContain("Name,Score");
  });
});

