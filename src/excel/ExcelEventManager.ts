export class ExcelEventManager {
  private debounceTimer: ReturnType<typeof setTimeout> | null = null;
  private readonly debounceMs: number;
  private handlers: Map<string, (event: Excel.WorksheetChangedEventArgs) => void> = new Map();

  constructor(debounceMs: number = 600) {
    this.debounceMs = debounceMs;
  }

  // Separate event wiring from business logic
  async onCellChanged(
    handler: (value: string, address: string) => void
  ): Promise<void> {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      sheet.onChanged.add((event: Excel.WorksheetChangedEventArgs) => {
        const value = event.details?.valueAfter as string;
        if (!value) return;

        // Espera 600ms después del último cambio antes de llamar a Claude
        if (this.debounceTimer) clearTimeout(this.debounceTimer);
        this.debounceTimer = setTimeout(() => {
          handler(value, event.address);
        }, this.debounceMs);
      });

      await context.sync();
    });
  }

  async onSelectionChanged(
    handler: (address: string) => void
  ): Promise<void> {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      sheet.onSelectionChanged.add((event: Excel.WorksheetSelectionChangedEventArgs) => {
        handler(event.address);
      });

      await context.sync();
    });
  } 
}
