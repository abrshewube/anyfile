# AnyFile Excel Module v0.2.0 – Specification

## Purpose
- Deliver the first pluggable handler for spreadsheet formats within the AnyFile ecosystem.
- Provide read/write support for XLSX (priority) with an abstraction ready for legacy XLS (via adapters).
- Establish uniform metadata, worksheet iteration, and conversion hooks that align with the core contract.

## Scope (v0.2.0)
- Detect Excel files by MIME, extension, and magic bytes.
- Read workbook metadata (sheets, dimensions, named ranges).
- Stream worksheets into typed row objects.
- Write workbooks back to disk/buffers with basic cell typing.
- Register the Excel handler with `@anyfile/core` via side effects.

## Dependencies & Tech
- Runtime: TypeScript, Node.js 18+ (Deno/browser support planned later).
- Libraries:
  - `xlsx` (SheetJS) as a baseline parser/writer.
  - `@anyfile/core` for handler registration and shared types.
- Build/Test: inherits root setup (tsup, Vitest).

## Handler Contract
```ts
declare module "@anyfile/core" {
  interface ExcelReadOptions {
    sheet?: string;
    headerRow?: number;
  }

  interface ExcelWriteOptions {
    sheetName?: string;
    cellStyles?: boolean;
  }

  interface ExcelFileData {
    workbook: Workbook;
    worksheets: WorksheetDescriptor[];
    toJSON(): Record<string, unknown>;
  }
}
```

### Registry Integration
- `registerFileType({ type: "excel", extensions: ["xls", "xlsx"], detect, open })`.
- Detection strategy:
  - Check file signature (`PK\x03\x04`) with `[Content_Types].xml` entry for XLSX.
  - Fallback to extension and MIME for quick paths.

## Public API (`@anyfile/excel`)
```ts
import "@anyfile/excel"; // auto-register

import { Excel } from "@anyfile/excel";

const workbook = await Excel.open("./report.xlsx");
const totals = await workbook.readSheet("Totals");

// Sheet helpers
const sheets = workbook.getSheetNames();
const metadata = workbook.getMetadata();
workbook.addSheetFromCSV("Imported", "Col1,Col2\n1,2\n3,4");

// Cell helpers
const cell = workbook.getCell("Totals", 2, 3);
workbook.setCell("Totals", 5, 1, "Grand Total", {
  style: {
    bold: true,
    fontColor: "#0B5FFF",
    backgroundColor: "#E8F1FF",
    numberFormat: "$#,##0.00",
  },
});

// Formulas
Excel.registerFormula("DOUBLE", (value) => Number(value) * 2);
workbook.setCell("Totals", 10, 2, null, { formula: "A10+B10" });
const result = workbook.evaluateCell("Totals", 10, 2);
console.log(result.evaluatedValue);

const report = workbook.evaluateAll({ ignoreCircular: true });
console.log(report.circular.length);

const circular = workbook.findCircularReferences();
if (circular.length > 0) {
  console.warn("Circular formula detected", circular);
}

const summary = workbook.getFormulaSummary();
console.log(summary.customFormulas);

// Asset discovery (charts/images/macros)
const charts = await workbook.getCharts();
const images = await workbook.getImages();
const macros = await workbook.listMacros();

// CSV export
const csv = workbook.toCSV("Totals");

await workbook.write("./report.xlsx", workbook);
```

### Exports
- `Excel.open(source, options?)`
- `Excel.register()` (re-exports core register for advanced customization)
- `Excel.detect(source)`
- Types: `ExcelWorkbook`, `ExcelWorksheet`, `ExcelCell`, `ExcelReadOptions`, `ExcelWriteOptions`

## File Operations
- **Read**
  - Loads workbook into memory.
  - Exposes sheet listing plus helpers:
    - `listSheets(): WorksheetDescriptor[]`
    - `readSheet(nameOrIndex, options?): AsyncIterable<RowObject>`
- **Write**
  - Accepts mutated workbook or row collections.
  - Supports streaming writes for large datasets (future).
- **Convert**
  - Integrates with `AnyFile.convert("csv")` by delegating to core conversions roadmap (placeholder hook returning `NotImplementedError` in v0.2.0).

## Error Handling
- Typed error classes: `ExcelFileError`, `ExcelParseError`, `ExcelWriteError`.
- Surface actionable context (sheet name, cell coordinates).

## Testing Strategy
- Unit tests leveraging fixture XLSX files (tiny in-memory).
- Property-based tests for row iteration (optional).
- Integration test verifying registry auto-registration with core.

## Roadmap Alignment
- v0.2.0: XLSX read/write, handler registration, metadata.
- v0.3.0: Advanced styling, XLS support.
- v0.4.0: Streaming, CSV conversion.

## Documentation TODO
- Usage examples for Node and bundlers.
- Troubleshooting guide for large workbooks.
- Performance tips (streaming vs in-memory).

