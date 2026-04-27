import { SheetContext } from "../types";

export class ExcelReader {
  // Single responsibility: only reads from Excel
  async getUsedRange(): Promise<SheetContext> {
    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRange();
      usedRange.load(["values", "formulas", "address", "rowCount", "columnCount"]);
      await context.sync();

      return {
        address: usedRange.address,
        values: usedRange.values,
        formulas: usedRange.formulas,
        rowCount: usedRange.rowCount,
        columnCount: usedRange.columnCount,
      };
    });
  }

  async getSelectedCell(): Promise<SheetContext> {
    return Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["values", "formulas", "address", "rowCount", "columnCount"]);
      await context.sync();

      return {
        address: range.address,
        values: range.values,
        formulas: range.formulas,
        rowCount: range.rowCount,
        columnCount: range.columnCount,
      };
    });
  }
}
