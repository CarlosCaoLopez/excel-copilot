import { AutoCompleteService } from "../services/AutoCompleteService";
import { ExcelEventManager } from "../excel/ExcelEventManager";
import { Suggestion } from "../types";

export class TaskPaneController {
  private currentSuggestion: Suggestion | null = null;
  private isLoading: boolean = false;

  constructor(
    private autoCompleteService: AutoCompleteService,
    private eventManager: ExcelEventManager
  ) {}

  // ─── Bootstrap ────────────────────────────────────────────────────────────

  async init(): Promise<void> {
    this.bindStaticElements();
    await this.registerExcelEvents();
  }

  // ─── Event Registration ───────────────────────────────────────────────────

  private async registerExcelEvents(): Promise<void> {
    await this.eventManager.onCellChanged(async (value, address) => {
      await this.handleCellChanged(value, address);
    });
  }

  // ─── Excel Event Handlers ─────────────────────────────────────────────────

  private async handleCellChanged(value: string, address: string): Promise<void> {
    if (!value || value.length < 2) {
      this.clearSuggestion();
      return;
    }

    this.setLoading(true);

    try {
      const suggestion = await this.autoCompleteService.suggest(value);
      this.currentSuggestion = suggestion;
      this.renderSuggestion(suggestion, address);
    } catch (error) {
      this.renderError("Could not fetch suggestion. Please try again.");
    } finally {
      this.setLoading(false);
    }
  }

  // ─── UI Event Bindings ────────────────────────────────────────────────────

  private bindStaticElements(): void {
    // Accept button
    document.getElementById("btn-accept")?.addEventListener("click", () => this.handleAccept());

    // Dismiss button
    document.getElementById("btn-dismiss")?.addEventListener("click", () => this.clearSuggestion());

    // Manual prompt input + submit
    document
      .getElementById("btn-submit-prompt")
      ?.addEventListener("click", () => this.handleManualPrompt());

    document.getElementById("input-prompt")?.addEventListener("keydown", (e: KeyboardEvent) => {
      if (e.key === "Enter") this.handleManualPrompt();
    });
  }

  // ─── UI Action Handlers ───────────────────────────────────────────────────

  private async handleAccept(): Promise<void> {
    if (!this.currentSuggestion) return;

    this.setLoading(true);

    try {
      await this.autoCompleteService.acceptSuggestion(this.currentSuggestion);
      this.clearSuggestion();
    } catch (error) {
      this.renderError("Could not write to cell.");
    } finally {
      this.setLoading(false);
    }
  }

  private async handleManualPrompt(): Promise<void> {
    const input = document.getElementById("input-prompt") as HTMLInputElement;
    const value = input?.value?.trim();

    if (!value) return;

    this.setLoading(true);

    try {
      const suggestion = await this.autoCompleteService.suggest(value);
      this.currentSuggestion = suggestion;
      this.renderSuggestion(suggestion, "manual");
    } catch (error) {
      this.renderError("Could not process prompt.");
    } finally {
      this.setLoading(false);
    }
  }

  // ─── Render Helpers ───────────────────────────────────────────────────────

  private renderSuggestion(suggestion: Suggestion, address: string): void {
    const box = document.getElementById("suggestion-box");
    const addressLabel = document.getElementById("suggestion-address");
    const actions = document.getElementById("suggestion-actions");

    if (box) box.innerText = suggestion.text;
    if (addressLabel) addressLabel.innerText = address !== "manual" ? `Cell: ${address}` : "";
    if (actions) actions.style.display = "flex";

    this.clearError();
  }

  private clearSuggestion(): void {
    this.currentSuggestion = null;

    const box = document.getElementById("suggestion-box");
    const addressLabel = document.getElementById("suggestion-address");
    const actions = document.getElementById("suggestion-actions");

    if (box) box.innerText = "";
    if (addressLabel) addressLabel.innerText = "";
    if (actions) actions.style.display = "none";
  }

  private setLoading(state: boolean): void {
    this.isLoading = state;

    const spinner = document.getElementById("spinner");
    const btnAccept = document.getElementById("btn-accept") as HTMLButtonElement;
    const btnSubmit = document.getElementById("btn-submit-prompt") as HTMLButtonElement;

    if (spinner) spinner.style.display = state ? "block" : "none";
    if (btnAccept) btnAccept.disabled = state;
    if (btnSubmit) btnSubmit.disabled = state;
  }

  private renderError(message: string): void {
    const errorEl = document.getElementById("error-message");
    if (errorEl) {
      errorEl.innerText = message;
      errorEl.style.display = "block";
    }
  }

  private clearError(): void {
    const errorEl = document.getElementById("error-message");
    if (errorEl) {
      errorEl.innerText = "";
      errorEl.style.display = "none";
    }
  }
}
