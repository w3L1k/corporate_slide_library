import type { SlideLibraryItem } from "@slide-library/shared";
import { formatUpdatedDate } from "./formatDate";
import { PreviewImage } from "./PreviewImage";
import { StatusBadge } from "./StatusBadge";

interface SlideCardProps {
  item: SlideLibraryItem;
  previewUrl: string;
  inserting: boolean;
  insertionBlocked: boolean;
  onOpen(): void;
  onInsert(): void;
}

export function SlideCard({
  item,
  previewUrl,
  inserting,
  insertionBlocked,
  onOpen,
  onInsert
}: SlideCardProps) {
  return (
    <article className="slide-card" aria-labelledby={`slide-${item.id}-title`}>
      <button
        className="slide-card__preview"
        type="button"
        onClick={onOpen}
        aria-label={`View details for ${item.title}`}
      >
        <PreviewImage src={previewUrl} title={item.title} />
        <span className="slide-card__preview-action" aria-hidden="true">
          Enlarge
        </span>
      </button>
      <div className="slide-card__body">
        <div className="slide-card__heading">
          <button
            id={`slide-${item.id}-title`}
            className="slide-card__title"
            type="button"
            onClick={onOpen}
          >
            {item.title}
          </button>
          <StatusBadge status={item.status} />
        </div>
        <p className="slide-card__category">{item.category}</p>
        <div className="slide-card__footer">
          <span className="slide-card__updated">Updated {formatUpdatedDate(item.updatedAt)}</span>
          <button
            className="button button--compact button--primary"
            type="button"
            onClick={onInsert}
            disabled={insertionBlocked}
            aria-label={`${inserting ? "Inserting" : "Insert"} ${item.title}`}
          >
            {inserting ? (
              <>
                <span className="spinner" aria-hidden="true" />
                Inserting…
              </>
            ) : (
              <>
                Insert
                <svg viewBox="0 0 16 16" aria-hidden="true">
                  <path d="M3 8h9M9 4l4 4-4 4" />
                </svg>
              </>
            )}
          </button>
        </div>
      </div>
    </article>
  );
}
