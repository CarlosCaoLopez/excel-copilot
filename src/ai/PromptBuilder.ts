import { SheetContext } from "../types";

export class PromptBuilder {
  // Isolated so you can iterate on prompts without touching AI or Excel logic
  buildAutoCompletePrompt(sheetContext: SheetContext, userInput: string): string {
    const serialized = this.serializeSheet(sheetContext);

    return `
      You are an Excel copilot. Based on the spreadsheet data and what the user
      is currently typing, suggest a formula or value to complete the cell.
      Respond with ONLY the raw completion — no explanation, no markdown.

      Sheet data:
      ${serialized}

      User is typing: "${userInput}"
    `.trim();
  }

  private serializeSheet(ctx: SheetContext): string {
    return ctx.values
      .map((row, i) =>
        row
          .map((cell, j) => {
            const formula = ctx.formulas[i][j];
            return formula !== cell ? formula : cell;
          })
          .join("\t")
      )
      .join("\n");
  }
}
