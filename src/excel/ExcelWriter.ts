export class ExcelWriter {
  private ghostAddress: string | null = null;

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

  // Escribe sugerencia en celda adyacente con formato gris
  async writeGhostText(anchorAddress: string, suggestion: string): Promise<void> {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const anchor = sheet.getRange(anchorAddress);
      anchor.load("columnIndex");
      await context.sync();

      // Celda a la derecha
      const ghostRange = anchor.getOffsetRange(0, 1);
      ghostRange.values = [[suggestion]];
      ghostRange.format.font.color = "#AAAAAA";
      ghostRange.format.font.italic = true;

      ghostRange.load("address");
      await context.sync();

      this.ghostAddress = ghostRange.address;
    });
  }

  // Acepta el ghost text — lo copia a la celda original y borra el ghost
  async acceptGhostText(targetAddress: string): Promise<void> {
    if (!this.ghostAddress) return;

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const ghostRange = sheet.getRange(this.ghostAddress!);
      ghostRange.load("values");
      await context.sync();

      const value = ghostRange.values[0][0];

      // Escribe en celda target
      const targetRange = sheet.getRange(targetAddress);
      targetRange.values = [[value]];

      // Borra el ghost
      ghostRange.clear();
      await context.sync();

      this.ghostAddress = null;
    });
  }

  // Descarta el ghost text sin aceptar
  async clearGhostText(): Promise<void> {
    if (!this.ghostAddress) return;

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const ghostRange = sheet.getRange(this.ghostAddress!);
      ghostRange.clear();
      await context.sync();
      this.ghostAddress = null;
    });
  }

  hasGhost(): boolean {
    return this.ghostAddress !== null;
  }
}