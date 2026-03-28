import type { ReactNode } from "react";
import type { ChatMessage } from "../types";

export function MessageThread({
  messages,
  confirmation,
}: {
  messages: ChatMessage[];
  confirmation?: ReactNode;
}) {
  return (
    <section className="message-thread" aria-label={"\u6d88\u606f\u7ebf\u7a0b"}>
      {messages.map((message) => (
        <article key={message.id} className={`message message-${message.role}`}>
          {message.content}
        </article>
      ))}
      {confirmation}
    </section>
  );
}
