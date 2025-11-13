# @anyfile/excel

Excel handler for the AnyFile ecosystem. This package plugs into `@anyfile/core` and registers spreadsheet support (XLSX, XLSM, XLSB, XLS) with a consistent API.

## Installation

```bash
pnpm add @anyfile/core @anyfile/excel
# or
npm install @anyfile/core @anyfile/excel
```

## Usage

Importing the package registers the handler automatically:

```ts
import "@anyfile/excel";
import { AnyFile } from "@anyfile/core";

const file = await AnyFile.open("./report.xlsx");
const workbook = await file.read();
const rows = await workbook.readSheet("Sheet1");

await file.write("./report-final.xlsx", workbook);
```

You can also use the module-specific helper:

```ts
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

// CSV export
const csv = workbook.toCSV("Totals");

await workbook.write("./report.xlsx", workbook);
```

## Features

- Detects Excel files by extension or signature.
- Reads workbook metadata, sheet descriptors, and row data.
- Sheet management helpers (list/add/delete/import CSV tabs).
- Cell-level read/write helpers with 1-based coordinates and styling options (font/bold/underline, fills, alignment, number formats).
- Writes workbooks to disk with automatic format detection.
- CSV export/import and `convert("csv")` integration with `AnyFile`.
- Seamlessly integrates with the `AnyFile` registry and conversion roadmap.

## Status

- **v0.2.0**: XLSX read/write, registry integration, sheet & cell helpers, CSV export/import.
- Upcoming releases will add styling, formula evaluation, charts/images, and PDF conversion.

