import type { ExcelCommandPreview } from '../types/bridge';

type ConfirmationCardProps = {
  preview: ExcelCommandPreview;
  isBusy: boolean;
  onConfirm: () => void;
  onCancel: () => void;
};

export function ConfirmationCard({ preview, isBusy, onConfirm, onCancel }: ConfirmationCardProps) {
  return (
    <article className="confirmation-card" aria-label="Confirm Excel action">
      <div className="confirmation-card__eyebrow">Pending write action</div>
      <h2 className="confirmation-card__title">Confirm Excel action</h2>
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
          Cancel
        </button>
        <button type="button" className="send-button" onClick={onConfirm} disabled={isBusy}>
          Confirm
        </button>
      </div>
    </article>
  );
}
