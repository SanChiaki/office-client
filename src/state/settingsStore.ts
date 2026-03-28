import { getJson, setJson } from "./localStorageAdapter";

const SETTINGS_KEY = "oa:settings";

export interface SettingsState {
  apiKey: string;
  model: string;
}

export function createSettingsStore() {
  return {
    load(): SettingsState {
      return getJson<SettingsState>(SETTINGS_KEY, {
        apiKey: "",
        model: "gpt-4.1-mini",
      });
    },
    save(value: SettingsState) {
      setJson(SETTINGS_KEY, value);
    },
  };
}
