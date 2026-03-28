import type { SelectionContext, ChatMessage } from "../types";

export function buildPrompt(input: string, messages: ChatMessage[], selection: SelectionContext | null) {
  return {
    input,
    messages,
    selection,
  };
}
