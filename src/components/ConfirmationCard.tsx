interface ConfirmationCardProps {
  summary: string;
  onConfirm(): void;
  onCancel(): void;
}

export function ConfirmationCard({ summary, onConfirm, onCancel }: ConfirmationCardProps) {
  return (
    <article className="confirmation-card">
      <p>{summary}</p>
      <div className="confirmation-actions">
        <button type="button" onClick={onConfirm}>
          {"\u786e\u8ba4"}
        </button>
        <button type="button" onClick={onCancel}>
          {"\u53d6\u6d88"}
        </button>
      </div>
    </article>
  );
}
