declare module "xlsx-calc" {
  import type { WorkBook } from "xlsx";

  function XLSXCalc(workbook: WorkBook): void;
  export = XLSXCalc;
}

