import { DEFAULT_BUSINESS_API_BASE_URL } from "../api/businessApiClient";
import { getJson, setJson } from "./localStorageAdapter";

const SETTINGS_KEY = "oa:settings";

export interface SettingsState {
  apiKey: string;
  baseUrl: string;
  model: string;
}

function normalizeSettings(value: Partial<SettingsState> | null | undefined): SettingsState {
  return {
    apiKey: typeof value?.apiKey === "string" ? value.apiKey : "",
    baseUrl:
      typeof value?.baseUrl === "string" && value.baseUrl.trim()
        ? value.baseUrl
        : DEFAULT_BUSINESS_API_BASE_URL,
    model: typeof value?.model === "string" && value.model ? value.model : "gpt-4.1-mini",
  };
}

export function createSettingsStore() {
  return {
    load(): SettingsState {
      return normalizeSettings(
        getJson<Partial<SettingsState>>(SETTINGS_KEY, {
          apiKey: "",
          baseUrl: DEFAULT_BUSINESS_API_BASE_URL,
          model: "gpt-4.1-mini",
        }),
      );
    },
    save(value: SettingsState) {
      setJson(SETTINGS_KEY, value);
    },
  };
}
