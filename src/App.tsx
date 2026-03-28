import { useEffect, useMemo, useState } from "react";
import { Composer } from "./components/Composer";
import { ConfirmationCard } from "./components/ConfirmationCard";
import { MessageThread } from "./components/MessageThread";
import { SettingsDialog } from "./components/SettingsDialog";
import { SelectionBadge } from "./components/SelectionBadge";
import { SessionSidebar } from "./components/SessionSidebar";
import { decideRoute } from "./agent/agentOrchestrator";
import { classifyAction, createExcelAdapter } from "./excel/excelAdapter";
import { subscribeToSelectionChanges } from "./excel/selectionContextService";
import { createSessionStore } from "./state/sessionStore";
import { createSettingsStore } from "./state/settingsStore";
import type { ExcelAction } from "./excel/excelAdapter";
import type { SettingsState } from "./state/settingsStore";
import type { ChatMessage, SelectionContext } from "./types";

const initialMessages: ChatMessage[] = [
  {
    id: "assistant-welcome",
    role: "assistant",
    content: "\u4f60\u597d\uff0c\u6211\u662f OfficeAgent\u3002",
  },
];

export default function App() {
  const sessionStore = useMemo(() => createSessionStore(), []);
  const settingsStore = useMemo(() => createSettingsStore(), []);
  const excelAdapter = useMemo(() => createExcelAdapter(), []);
  const [{ sessions, activeSessionId }, setSessionState] = useState(sessionStore.getState());
  const [settings, setSettings] = useState(settingsStore.load());
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [draft, setDraft] = useState("");
  const [messages, setMessages] = useState<ChatMessage[]>(initialMessages);
  const [selection, setSelection] = useState<SelectionContext | null>(null);
  const [pendingConfirmation, setPendingConfirmation] = useState<ExcelAction | null>(null);

  const activeSession = sessions.find((session) => session.id === activeSessionId) ?? null;

  function refreshSessions() {
    setSessionState(sessionStore.getState());
  }

  function syncMessages(nextMessages: ChatMessage[]) {
    setMessages(nextMessages);

    if (activeSessionId) {
      sessionStore.replaceMessages(activeSessionId, nextMessages);
      refreshSessions();
    }
  }

  useEffect(() => {
    if (!sessionStore.getState().sessions.length) {
      sessionStore.createSession();
      refreshSessions();
    }
  }, [sessionStore]);

  useEffect(() => {
    setMessages(activeSession?.messages.length ? activeSession.messages : initialMessages);
  }, [activeSessionId, activeSession]);

  useEffect(() => {
    return subscribeToSelectionChanges((nextSelection) => {
      setSelection(nextSelection);
    });
  }, []);

  function handleCreateSession() {
    sessionStore.createSession();
    setDraft("");
    refreshSessions();
  }

  function handleSelectSession(id: string) {
    sessionStore.setActiveSession(id);
    refreshSessions();
  }

  function handleDeleteSession(id: string) {
    sessionStore.deleteSession(id);
    refreshSessions();
  }

  function handleOpenSettings() {
    setSettings(settingsStore.load());
    setIsSettingsOpen(true);
  }

  function handleSaveSettings(nextSettings: SettingsState) {
    settingsStore.save(nextSettings);
    setSettings(nextSettings);
    setIsSettingsOpen(false);
  }

  async function executeExcelAction(action: ExcelAction) {
    try {
      await excelAdapter.run(action);
    } finally {
      setPendingConfirmation(null);
    }
  }

  function queueAction(action: ExcelAction) {
    if (classifyAction(action).requiresConfirmation) {
      setPendingConfirmation(action);
      return;
    }

    void executeExcelAction(action);
  }

  function handleSubmit() {
    const content = draft.trim();
    if (!content) {
      return;
    }

    const route = decideRoute(content);
    if (route.mode !== "chat") {
      return;
    }

    const userMessage: ChatMessage = {
      id: crypto.randomUUID(),
      role: "user",
      content,
    };

    const nextMessages = activeSession?.messages.length ? [...messages, userMessage] : [userMessage];
    syncMessages(nextMessages);
    setDraft("");
  }

  return (
    <main className="layout">
      <SessionSidebar
        sessions={sessions}
        activeSessionId={activeSessionId ?? undefined}
        onCreateSession={handleCreateSession}
        onSelectSession={handleSelectSession}
        onDeleteSession={handleDeleteSession}
      />
      <section className="chat-panel">
        <header className="chat-header">
          <h1>OfficeAgent</h1>
          <button type="button" className="chat-header-action" onClick={handleOpenSettings}>
            {"\u8bbe\u7f6e"}
          </button>
        </header>
        {isSettingsOpen ? (
          <SettingsDialog
            initialValue={settings}
            onSave={handleSaveSettings}
            onClose={() => setIsSettingsOpen(false)}
          />
        ) : null}
        <MessageThread
          messages={messages}
          confirmation={
            pendingConfirmation ? (
              <ConfirmationCard
                summary={`Prepare to run ${pendingConfirmation.type}`}
                onConfirm={() => {
                  void executeExcelAction(pendingConfirmation);
                }}
                onCancel={() => setPendingConfirmation(null)}
              />
            ) : null
          }
        />
        <SelectionBadge selection={selection} />
        <Composer value={draft} onChange={setDraft} onSubmit={handleSubmit} />
      </section>
    </main>
  );
}
