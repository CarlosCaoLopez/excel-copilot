import Anthropic from "@anthropic-ai/sdk";
import { AutoCompleteOptions, Suggestion } from "../types";

export class ClaudeClient {
  private client: Anthropic;

  constructor(apiKey: string) {
    this.client = new Anthropic({ apiKey, dangerouslyAllowBrowser: true });
  }

  // Only responsibility: talk to the Claude API
  async complete(prompt: string, options: AutoCompleteOptions = {}): Promise<Suggestion> {
    const response = await this.client.messages.create({
      model: "claude-sonnet-4-20250514",
      max_tokens: options.maxTokens ?? 256,
      messages: [{ role: "user", content: prompt }],
    });

    const text = response.content[0].type === "text" ? response.content[0].text.trim() : "";

    return { text };
  }
}
