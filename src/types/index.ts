export interface SheetContext {
  address: string; // Active cell
  values: unknown[][]; // Matrix with all raw data
  formulas: string[][]; // Matrix with all formulas
  rowCount: number;
  columnCount: number;
}

export interface Suggestion {
  text: string;
  confidence?: number;
}

export interface AutoCompleteOptions {
  maxTokens?: number;
  temperature?: number;
}
