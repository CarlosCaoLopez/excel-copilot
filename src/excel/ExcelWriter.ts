export class ExcelWriter {
  // Single responsibility: only writes to Excel
  async writeToSelectedCell(value: string): Promise<void> {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.values = [[value]];
      await context.sync();
    });
  }

  async writeToAddress(address: string, value: string): Promise<void> {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(address);
      range.values = [[value]];
      await context.sync();
    });
  }
}
