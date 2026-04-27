import { ExcelReader } from "../excel/ExcelReader";
import { ExcelWriter } from "../excel/ExcelWriter";
import { ExcelEventManager } from "../excel/ExcelEventManager";
import { ClaudeClient } from "../ai/ClaudeClient";
import { PromptBuilder } from "../ai/PromptBuilder";
import { AutoCompleteService } from "../services/AutoCompleteService";
import { TaskPaneController } from "../ui/TaskPaneController";



const apiKey = process.env.API_KEY;
if (!apiKey) throw new Error("API_KEY is not defined in .env");
Office.onReady(() => {
  // Wire up dependencies here — one place, easy to swap
  const reader = new ExcelReader();
  const writer = new ExcelWriter();
  const events = new ExcelEventManager();
  const claude = new ClaudeClient(apiKey);
  const promptBuilder = new PromptBuilder();
  const autoComplete = new AutoCompleteService(reader, writer, claude, promptBuilder);
  const ui = new TaskPaneController(autoComplete, events);

  ui.init();
});
