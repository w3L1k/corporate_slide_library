import type { SlideLibraryItem } from "@slide-library/shared";
import { formatCategory } from "./formatCategory";
import { formatUpdatedDate } from "./formatDate";
import { PreviewImage } from "./PreviewImage";
import { StatusBadge } from "./StatusBadge";

interface SlideCardProps {
  item: SlideLibraryItem;
  previewUrl: string;
  inserting: boolean;
  insertionBlocked: boolean;
  favorite: boolean;
  selected: boolean;
  onOpen(): void;
  onInsert(): void;
  onToggleFavorite(): void;
  onToggleSelection(): void;
}

export function SlideCard({
  item,
  previewUrl,
  inserting,
  insertionBlocked,
  favorite,
  selected,
  onOpen,
  onInsert,
  onToggleFavorite,
  onToggleSelection
}: SlideCardProps) {
  return (
    <article
      className={`slide-card ${selected ? "slide-card--selected" : ""}`}
      aria-labelledby={`slide-${item.id}-title`}
    >
      <button
        className="slide-card__select"
        type="button"
        onClick={onToggleSelection}
        aria-label={`${selected ? "Снять выбор" : "Выбрать"}: ${item.title}`}
        aria-pressed={selected}
        title={selected ? "Снять выбор" : "Выбрать слайд"}
      >
        <svg viewBox="0 0 20 20" aria-hidden="true">
          <rect x="3" y="3" width="14" height="14" rx="4" />
          {selected ? <path d="M6.5 10l2.2 2.3 4.8-5" /> : null}
        </svg>
      </button>
      <button
        className="slide-card__favorite"
        type="button"
        onClick={onToggleFavorite}
        aria-label={`${
          favorite ? "Убрать из избранного" : "Добавить в избранное"
        }: ${item.title}`}
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
        aria-label={`Открыть сведения: ${item.title}`}
      >
        <PreviewImage src={previewUrl} title={item.title} />
        <span className="slide-card__preview-action" aria-hidden="true">
          Увеличить
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
        <p className="slide-card__category">{formatCategory(item.category)}</p>
        <div className="slide-card__footer">
          <span className="slide-card__updated">
            Обновлено {formatUpdatedDate(item.updatedAt)}
          </span>
          <button
            className="button button--compact button--primary"
            type="button"
            onClick={onInsert}
            disabled={insertionBlocked}
            aria-label={`${
              inserting ? "Добавляется в PowerPoint" : "Добавить в PowerPoint"
            }: ${item.title}`}
          >
            {inserting ? (
              <>
                <span className="spinner" aria-hidden="true" />
                Добавляем…
              </>
            ) : (
              <>
                Добавить
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
