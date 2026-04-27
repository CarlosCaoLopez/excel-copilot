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

  async acceptSuggestion(suggestion: Suggestion): Promise<void> {
    await this.writer.writeToSelectedCell(suggestion.text);
  }
}
