declare module "xlsx-calc" {
  import type { WorkBook } from "xlsx";

  interface XLSXCalc {
    (workbook: WorkBook, options?: unknown): void;
    set_fx(name: string, fn: (...args: unknown[]) => unknown): void;
    import_functions(
      functions: Record<string, (...args: unknown[]) => unknown>,
      options?: unknown
    ): void;
    localizeFunctions(dictionary: Record<string, string>): void;
  }

  const calc: XLSXCalc;
  export = calc;
}

