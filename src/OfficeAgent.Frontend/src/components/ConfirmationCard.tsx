import type { ExcelCommandPreview } from '../types/bridge';

type ConfirmationCardProps = {
  preview: ExcelCommandPreview;
  isBusy: boolean;
  ariaLabel: string;
  eyebrow: string;
  title: string;
  cancelLabel: string;
  confirmLabel: string;
  onConfirm: () => void;
  onCancel: () => void;
};

export function ConfirmationCard({
  preview,
  isBusy,
  ariaLabel,
  eyebrow,
  title,
  cancelLabel,
  confirmLabel,
  onConfirm,
  onCancel,
}: ConfirmationCardProps) {
  return (
    <article className="confirmation-card" aria-label={ariaLabel}>
      <div className="confirmation-card__eyebrow">{eyebrow}</div>
      <h2 className="confirmation-card__title">{title}</h2>
      <p className="confirmation-card__summary">{preview.title}</p>
      <p className="confirmation-card__summary">{preview.summary}</p>
      {preview.details.length > 0 ? (
        <ul className="confirmation-card__details">
          {preview.details.map((detail) => (
            <li key={detail}>{detail}</li>
          ))}
        </ul>
      ) : null}
      <div className="confirmation-card__actions">
        <button type="button" className="ghost-button" onClick={onCancel} disabled={isBusy}>
          {cancelLabel}
        </button>
        <button type="button" className="send-button" onClick={onConfirm} disabled={isBusy}>
          {confirmLabel}
        </button>
      </div>
    </article>
  );
}
