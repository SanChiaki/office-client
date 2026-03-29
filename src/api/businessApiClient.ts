export const DEFAULT_BUSINESS_API_BASE_URL = "https://api.example.com";

export function normalizeBaseUrl(baseUrl: string) {
  const trimmedBaseUrl = baseUrl.trim();
  const resolvedBaseUrl = trimmedBaseUrl || DEFAULT_BUSINESS_API_BASE_URL;

  return resolvedBaseUrl.replace(/\/+$/, "");
}

export async function uploadData(apiKey: string, baseUrl: string, payload: unknown) {
  const response = await fetch(`${normalizeBaseUrl(baseUrl)}/upload_data_api`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    throw new Error(`Request failed with status ${response.status}`);
  }

  return response.json();
}
