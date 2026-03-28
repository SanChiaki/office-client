import type { CommandEnvelope } from "../agent/commandSchema";

export async function requestCommandEnvelope(apiKey: string, payload: unknown): Promise<CommandEnvelope> {
  const response = await fetch("https://api.example.com/agent", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  });

  return response.json();
}
