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

function createNumericWorkbookBuffer() {
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.aoa_to_sheet([
    ["Price", "Qty", "Total", "Circular"],
    [10, 2, null, null],
    [5, 4, null, null],
    [null, null, null, null],
  ]);

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
    file.setCell("Sheet1", 4, 3, 0.85, {
      style: {
        bold: true,
        fontColor: "#FF0000",
        backgroundColor: "#FFFFAA",
        numberFormat: "0.00%",
      },
    });

    const updatedRows = await file.readSheet("Sheet1", { headerRow: 1 });
    expect(updatedRows).toEqual([
      { Name: "Alice", Score: 95, column_3: null },
      { Name: "Bob", Score: 87, column_3: null },
      { Name: "Charlie", Score: 72, column_3: 0.85 },
    ]);

    const styledCell = file.getCell("Sheet1", 4, 3);
    expect(styledCell?.style?.bold).toBe(true);
    expect(styledCell?.style?.fontColor).toBe("FF0000");
    expect(styledCell?.style?.backgroundColor).toBe("FFFFAA");
    expect(styledCell?.style?.numberFormat).toBe("0.00%");

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

  it("returns empty chart/image/macro metadata when not present", async () => {
    const buffer = createWorkbookBuffer();
    const file = await Excel.open(buffer);

    expect(await file.getCharts()).toEqual([]);
    expect(await file.getImages()).toEqual([]);
    expect(await file.listMacros()).toEqual([]);
  });

  it("evaluates formulas and detects circular references", async () => {
    const buffer = createNumericWorkbookBuffer();
    const file = await Excel.open(buffer);

    file.setCell("Sheet1", 2, 3, null, { formula: "A2*B2" });
    file.setCell("Sheet1", 3, 3, null, { formula: "A3*B3" });
    file.setCell("Sheet1", 4, 3, null, { formula: "SUM(C2:C3)" });

    const totalCell = file.getCell("Sheet1", 2, 3);
    expect(totalCell?.formula).toBe("A2*B2");

    const result = file.evaluateCell("Sheet1", 2, 3);
    expect(result.value).toBe(20);
    expect(result.evaluatedValue).toBe(20);
    expect(result.error).toBeUndefined();

    const evaluationSummary = file.evaluateAll();
    expect(evaluationSummary.evaluated.length).toBeGreaterThan(0);
    expect(evaluationSummary.circular.length).toBe(0);

    const sumCell = file.getCell("Sheet1", 4, 3);
    expect(sumCell?.value).toBe(40);

    file.setCell("Sheet1", 2, 4, null, { formula: "D3" });
    file.setCell("Sheet1", 3, 4, null, { formula: "D2" });

    const circularReport = file.evaluateAll({ ignoreCircular: true });
    expect(circularReport.circular.length).toBeGreaterThan(0);

    const circularCell = file.evaluateCell("Sheet1", 2, 4);
    expect(circularCell.error).toContain("Circular");

    const circular = file.findCircularReferences();
    expect(circular.length).toBeGreaterThan(0);
    const flattened = circular.flatMap((entry) => entry.path);
    expect(flattened).toContain("Sheet1!D2");
    expect(flattened).toContain("Sheet1!D3");

    const summary = file.getFormulaSummary();
    expect(summary.totalFormulas).toBeGreaterThanOrEqual(3);
    expect(summary.circularReferences).toBe(circular.length);
  });

  it("supports custom formula registration and localization", async () => {
    Excel.registerFormula("DOUBLE", (value: unknown) => Number(value) * 2);
    Excel.registerFormulas({
      TRIPLE: (value: unknown) => Number(value) * 3,
    });
    Excel.configureLocalization({ SUMA: "SUM" });

    const buffer = createNumericWorkbookBuffer();
    const file = await Excel.open(buffer);

    file.setCell("Sheet1", 2, 3, null, { formula: "DOUBLE(A2)" });
    file.setCell("Sheet1", 3, 3, null, { formula: "TRIPLE(A3)" });
    file.setCell("Sheet1", 5, 3, null, { formula: "SUMA(A2:A3)" });

    const report = file.evaluateAll();
    expect(report.evaluated.length).toBeGreaterThan(0);

    const doubleCell = file.getCell("Sheet1", 2, 3);
    const tripleCell = file.getCell("Sheet1", 3, 3);
    const localizedSum = file.getCell("Sheet1", 5, 3);

    expect(doubleCell?.value).toBe(20);
    expect(tripleCell?.value).toBe(15);
    expect(localizedSum?.value).toBe(15);

    const summary = file.getFormulaSummary();
    expect(summary.customFormulas).toEqual(
      expect.arrayContaining(["DOUBLE", "TRIPLE"])
    );
  });
});

