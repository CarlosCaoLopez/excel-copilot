import { ExcelReader } from "../excel/ExcelReader";
import { ExcelWriter } from "../excel/ExcelWriter";
import { ClaudeClient } from "../ai/ClaudeClient";
import { PromptBuilder } from "../ai/PromptBuilder";
import { Suggestion } from "../types";

export class AutoCompleteService {
  // Orchestrator: coordinates the other classes, owns no low-level logic itself
  constructor(
    private reader: ExcelReader,
    private writer: ExcelWriter,
    private claude: ClaudeClient,
    private promptBuilder: PromptBuilder
  ) {}

  async suggest(userInput: string): Promise<Suggestion> {
    const sheetContext = await this.reader.getUsedRange();
    const prompt = this.promptBuilder.buildAutoCompletePrompt(sheetContext, userInput);
    return this.claude.complete(prompt);
  }

  // Muestra ghost text en la celda adyacente
  async suggestInline(userInput: string, address: string): Promise<void> {
    this.lastAnchorAddress = address;

    const sheetContext = await this.reader.getUsedRange();
    const prompt = this.promptBuilder.buildAutoCompletePrompt(sheetContext, userInput);
    const suggestion = await this.claude.complete(prompt);

    await this.writer.writeGhostText(address, suggestion.text);
  }

  // Tab presionado — acepta si hay ghost
  async acceptInline(): Promise<boolean> {
    if (!this.writer.hasGhost() || !this.lastAnchorAddress) return false;

    await this.writer.acceptGhostText(this.lastAnchorAddress);
    this.lastAnchorAddress = null;
    return true;
  }

  // Escape — descarta ghost
  async dismissInline(): Promise<void> {
    await this.writer.clearGhostText();
    this.lastAnchorAddress = null;
  }

  async acceptSuggestion(suggestion: Suggestion): Promise<void> {
    await this.writer.writeToSelectedCell(suggestion.text);
  }
}
