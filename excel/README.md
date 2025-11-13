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

await workbook.write("./report.xlsx", workbook);
```

## Features

- Detects Excel files by extension or signature.
- Reads workbook metadata, sheet descriptors, and row data.
- Writes workbooks to disk with automatic format detection.
- Seamlessly integrates with the `AnyFile` registry and conversion pipeline stubs.

## Status

- **v0.2.0**: XLSX read/write, registry integration, workbook helpers.
- Upcoming releases will add streaming support, conversions (CSV), and styling.

