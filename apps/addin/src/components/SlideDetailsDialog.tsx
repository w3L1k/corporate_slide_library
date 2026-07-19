import { useEffect, useRef } from "react";
import { createPortal } from "react-dom";
import type { SlideLibraryItem } from "@slide-library/shared";
import { formatCategory } from "./formatCategory";
import { formatUpdatedDate } from "./formatDate";
import { PreviewImage } from "./PreviewImage";
import { StatusBadge } from "./StatusBadge";

const FOCUSABLE_SELECTOR =
  'button:not([disabled]), [href], input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])';

interface SlideDetailsDialogProps {
  item: SlideLibraryItem;
  previewUrl: string;
  inserting: boolean;
  insertionBlocked: boolean;
  onClose(): void;
  onInsert(): void;
}

export function SlideDetailsDialog({
  item,
  previewUrl,
  inserting,
  insertionBlocked,
  onClose,
  onInsert
}: SlideDetailsDialogProps) {
  const dialogRef = useRef<HTMLDivElement>(null);
  const closeButtonRef = useRef<HTMLButtonElement>(null);

  useEffect(() => {
    const previouslyFocused = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    const previousOverflow = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    closeButtonRef.current?.focus();

    return () => {
      document.body.style.overflow = previousOverflow;
      previouslyFocused?.focus();
    };
  }, []);

  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent): void => {
      if (event.key === "Escape") {
        event.preventDefault();
        onClose();
        return;
      }

      if (event.key !== "Tab" || !dialogRef.current) {
        return;
      }

      const focusableElements = Array.from(
        dialogRef.current.querySelectorAll<HTMLElement>(FOCUSABLE_SELECTOR)
      );
      const firstElement = focusableElements[0];
      const lastElement = focusableElements.at(-1);

      if (!firstElement || !lastElement) {
        event.preventDefault();
        return;
      }

      if (event.shiftKey && document.activeElement === firstElement) {
        event.preventDefault();
        lastElement.focus();
      } else if (!event.shiftKey && document.activeElement === lastElement) {
        event.preventDefault();
        firstElement.focus();
      }
    };

    document.addEventListener("keydown", handleKeyDown);
    return () => document.removeEventListener("keydown", handleKeyDown);
  }, [onClose]);

  return createPortal(
    <div
      className="dialog-backdrop"
      onMouseDown={(event) => {
        if (event.target === event.currentTarget) {
          onClose();
        }
      }}
    >
      <div
        ref={dialogRef}
        className="detail-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="detail-title"
        aria-describedby={item.description ? "detail-description" : undefined}
      >
        <header className="detail-dialog__header">
          <div>
            <span className="eyebrow">Сведения о слайде</span>
            <h2 id="detail-title">{item.title}</h2>
          </div>
          <button
            ref={closeButtonRef}
            className="icon-button"
            type="button"
            onClick={onClose}
            aria-label="Закрыть сведения о слайде"
          >
            <svg viewBox="0 0 20 20" aria-hidden="true">
              <path d="M5 5l10 10M15 5L5 15" />
            </svg>
          </button>
        </header>

        <div className="detail-dialog__content">
          <div className="detail-dialog__preview">
            <PreviewImage src={previewUrl} title={item.title} eager />
          </div>

          <section className="detail-dialog__info" aria-label="Метаданные слайда">
            <div className="detail-dialog__status-row">
              <StatusBadge status={item.status} />
              <span>Версия {item.version}</span>
            </div>
            {item.description ? <p id="detail-description">{item.description}</p> : null}

            <dl className="metadata-grid">
              <div>
                <dt>Категория</dt>
                <dd>{formatCategory(item.category)}</dd>
              </div>
              <div>
                <dt>Обновлено</dt>
                <dd>{formatUpdatedDate(item.updatedAt)}</dd>
              </div>
              {item.department ? (
                <div>
                  <dt>Подразделение</dt>
                  <dd>{item.department}</dd>
                </div>
              ) : null}
              {item.owner ? (
                <div>
                  <dt>Владелец контента</dt>
                  <dd>{item.owner}</dd>
                </div>
              ) : null}
              {item.author ? (
                <div>
                  <dt>Автор</dt>
                  <dd>{item.author}</dd>
                </div>
              ) : null}
              {item.language ? (
                <div>
                  <dt>Язык</dt>
                  <dd>{item.language}</dd>
                </div>
              ) : null}
            </dl>

            <div className="tag-list" aria-label="Теги">
              {item.tags.map((tag) => (
                <span key={tag}>{tag}</span>
              ))}
            </div>
          </section>
        </div>

        <footer className="detail-dialog__footer">
          <button className="button button--secondary" type="button" onClick={onClose}>
            Закрыть
          </button>
          <button
            className="button button--primary button--insert"
            type="button"
            onClick={onInsert}
            disabled={insertionBlocked}
          >
            {inserting ? (
              <>
                <span className="spinner" aria-hidden="true" />
                Добавляем…
              </>
            ) : (
              <>
                <svg viewBox="0 0 20 20" aria-hidden="true">
                  <rect x="3" y="4" width="14" height="11" rx="1.5" />
                  <path d="M10 7v5M7.5 9.5h5M7 17h6" />
                </svg>
                Добавить слайд
              </>
            )}
          </button>
        </footer>
      </div>
    </div>,
    document.body
  );
}
