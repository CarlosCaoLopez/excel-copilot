export class ExcelEventManager {
  private handlers: Map<string, (event: Excel.WorksheetChangedEventArgs) => void> = new Map();

  // Separate event wiring from business logic
  async onCellChanged(handler: (value: string, address: string) => void): Promise<void> {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      sheet.onChanged.add((event: Excel.WorksheetChangedEventArgs) => {
        const value = event.details?.valueAfter as string;
        if (value) handler(value, event.address);
      });

      await context.sync();
    });
  }
}
