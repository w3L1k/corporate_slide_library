import type { SlideLibraryItem } from "@slide-library/shared";
import { formatUpdatedDate } from "./formatDate";
import { PreviewImage } from "./PreviewImage";
import { StatusBadge } from "./StatusBadge";

interface SlideCardProps {
  item: SlideLibraryItem;
  previewUrl: string;
  inserting: boolean;
  insertionBlocked: boolean;
  favorite: boolean;
  onOpen(): void;
  onInsert(): void;
  onToggleFavorite(): void;
}

export function SlideCard({
  item,
  previewUrl,
  inserting,
  insertionBlocked,
  favorite,
  onOpen,
  onInsert,
  onToggleFavorite
}: SlideCardProps) {
  return (
    <article className="slide-card" aria-labelledby={`slide-${item.id}-title`}>
      <button
        className="slide-card__favorite"
        type="button"
        onClick={onToggleFavorite}
        aria-label={`${favorite ? "Remove" : "Add"} ${item.title} ${
          favorite ? "from" : "to"
        } favorites`}
        aria-pressed={favorite}
        title={favorite ? "Убрать из избранного" : "Добавить в избранное"}
      >
        <svg viewBox="0 0 20 20" aria-hidden="true">
          <path d="M10 17s-6-3.7-6-8.3a3.5 3.5 0 016-2.3 3.5 3.5 0 016 2.3C16 13.3 10 17 10 17z" />
        </svg>
      </button>
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
