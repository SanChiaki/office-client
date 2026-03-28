import { getJson, setJson } from "./localStorageAdapter";

const SETTINGS_KEY = "oa:settings";

export interface SettingsState {
  apiKey: string;
  model: string;
}

function normalizeSettings(value: Partial<SettingsState> | null | undefined): SettingsState {
  return {
    apiKey: typeof value?.apiKey === "string" ? value.apiKey : "",
    model: typeof value?.model === "string" && value.model ? value.model : "gpt-4.1-mini",
  };
}

export function createSettingsStore() {
  return {
    load(): SettingsState {
      return normalizeSettings(
        getJson<Partial<SettingsState>>(SETTINGS_KEY, {
          apiKey: "",
          model: "gpt-4.1-mini",
        })
      );
    },
    save(value: SettingsState) {
      setJson(SETTINGS_KEY, value);
    },
  };
}
