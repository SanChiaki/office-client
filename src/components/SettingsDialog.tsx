import { useState } from "react";
import type { SettingsState } from "../state/settingsStore";

export function SettingsDialog({
  initialValue,
  onSave,
  onClose,
}: {
  initialValue: SettingsState;
  onSave(value: SettingsState): void;
  onClose(): void;
}) {
  const [apiKey, setApiKey] = useState(initialValue.apiKey);
  const [model, setModel] = useState(initialValue.model);

  return (
    <section className="settings-dialog" aria-label="Settings">
      <label>
        API Key
        <input value={apiKey} onChange={(event) => setApiKey(event.target.value)} />
      </label>
      <label>
        Model
        <input value={model} onChange={(event) => setModel(event.target.value)} />
      </label>
      <div className="settings-dialog-actions">
        <button type="button" onClick={onClose}>
          取消
        </button>
        <button type="button" onClick={() => onSave({ apiKey, model })}>
          保存设置
        </button>
      </div>
    </section>
  );
}
