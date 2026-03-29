import { useEffect, useState } from 'react';
import { nativeBridge } from './bridge/nativeBridge';

export function App() {
  const [bridgeStatus, setBridgeStatus] = useState('Connecting to native host...');

  useEffect(() => {
    let isActive = true;

    nativeBridge
      .ping()
      .then((result) => {
        if (!isActive) {
          return;
        }

        setBridgeStatus(`Connected to ${result.host} (${result.version})`);
      })
      .catch((error: Error) => {
        if (!isActive) {
          return;
        }

        setBridgeStatus(`Native bridge unavailable: ${error.message}`);
      });

    return () => {
      isActive = false;
    };
  }, []);

  return (
    <div className="app-shell">
      <aside className="sidebar" aria-label="Session sidebar placeholder">
        <div className="sidebar__title">Sessions</div>
        <div className="sidebar__empty">No sessions yet</div>
      </aside>

      <main className="workspace">
        <header className="chat-header" aria-label="Chat header">
          <div>
            <div className="eyebrow">Office Agent</div>
            <h1 className="title">Task pane shell</h1>
          </div>

          <button type="button" className="icon-button" aria-label="Settings">
            Settings
          </button>
        </header>

        <section className="selection-badge" aria-label="Selection badge placeholder" role="status">
          <div className="selection-badge__label">Bridge status</div>
          <div>{bridgeStatus}</div>
        </section>

        <section className="thread" aria-label="Message thread">
          <article className="message message--assistant">
            <p>Welcome to Office Agent. This shell is ready for the chat experience.</p>
          </article>
        </section>

        <footer className="composer" aria-label="Message composer">
          <textarea
            aria-label="Message composer"
            placeholder="Type a message..."
            rows={3}
          />
          <button type="button" className="send-button">
            Send
          </button>
        </footer>
      </main>
    </div>
  );
}

export default App;
