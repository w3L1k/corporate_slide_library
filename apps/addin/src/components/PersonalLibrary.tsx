import { useEffect, useMemo, useRef, useState, type FormEvent } from "react";
import type { PersonalAsset, PersonalAssetKind } from "@slide-library/shared";
import type { SlideLibraryApi } from "../services/api";
import {
  PowerPointUnavailableError,
  type PowerPointService
} from "../services/powerpoint";
import type { ToastMessage } from "./Toast";

const MAX_FILE_SIZE = 20 * 1024 * 1024;

const kindLabels: Record<PersonalAssetKind, string> = {
  presentation: "Презентации",
  photo: "Фотографии",
  illustration: "Иллюстрации",
  logo: "Логотипы"
};

const kindAccepts: Record<PersonalAssetKind, string> = {
  presentation:
    ".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation",
  photo: ".png,.jpg,.jpeg,.webp,image/png,image/jpeg,image/webp",
  illustration:
    ".png,.jpg,.jpeg,.webp,.svg,image/png,image/jpeg,image/webp,image/svg+xml",
  logo: ".png,.jpg,.jpeg,.webp,.svg,image/png,image/jpeg,image/webp,image/svg+xml"
};

const kindHints: Record<PersonalAssetKind, string> = {
  presentation: "PPTX до 20 МБ",
  photo: "PNG, JPEG или WebP до 20 МБ",
  illustration: "PNG, JPEG, WebP или безопасный SVG до 20 МБ",
  logo: "PNG, JPEG, WebP или безопасный SVG до 20 МБ"
};

const insertionSuccessMessages: Record<PersonalAssetKind, string> = {
  presentation: "Презентация добавлена в PowerPoint.",
  photo: "Фотография добавлена на текущий слайд.",
  illustration: "Иллюстрация добавлена на текущий слайд.",
  logo: "Логотип добавлен на текущий слайд."
};

interface PersonalLibraryProps {
  api: SlideLibraryApi;
  kind: PersonalAssetKind;
  powerPointService: PowerPointService;
  onNotify(message: ToastMessage): void;
}

const formatFileSize = (size: number): string => {
  if (size < 1024 * 1024) {
    return `${Math.max(1, Math.round(size / 1024))} КБ`;
  }
  return `${(size / (1024 * 1024)).toFixed(1)} МБ`;
};

const getErrorMessage = (error: unknown): string =>
  error instanceof Error && error.message.trim()
    ? error.message
    : "Не удалось выполнить операцию.";

export function PersonalLibrary({
  api,
  kind,
  powerPointService,
  onNotify
}: PersonalLibraryProps) {
  const [items, setItems] = useState<PersonalAsset[]>([]);
  const [loading, setLoading] = useState(true);
  const [uploading, setUploading] = useState(false);
  const [insertingId, setInsertingId] = useState<string | null>(null);
  const [deletingId, setDeletingId] = useState<string | null>(null);
  const [confirmingDeleteId, setConfirmingDeleteId] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [title, setTitle] = useState("");
  const [file, setFile] = useState<File | null>(null);
  const [refreshKey, setRefreshKey] = useState(0);
  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    const controller = new AbortController();

    void api
      .listPersonalAssets(controller.signal)
      .then((response) => {
        if (!controller.signal.aborted) {
          setItems(response.items);
          setLoading(false);
        }
      })
      .catch((requestError: unknown) => {
        if (!controller.signal.aborted) {
          setError(getErrorMessage(requestError));
          setLoading(false);
        }
      });

    return () => controller.abort();
  }, [api, refreshKey]);

  const visibleItems = useMemo(
    () => items.filter((item) => item.kind === kind),
    [items, kind]
  );

  const submitUpload = async (event: FormEvent<HTMLFormElement>): Promise<void> => {
    event.preventDefault();
    const normalizedTitle = title.trim();
    if (!normalizedTitle || !file) {
      onNotify({ kind: "info", text: "Укажите название и выберите файл." });
      return;
    }
    if (file.size > MAX_FILE_SIZE) {
      onNotify({ kind: "error", text: "Размер файла не должен превышать 20 МБ." });
      return;
    }

    setUploading(true);
    try {
      const item = await api.uploadPersonalAsset(kind, normalizedTitle, file);
      setItems((current) => [item, ...current]);
      setTitle("");
      setFile(null);
      if (fileInputRef.current) {
        fileInputRef.current.value = "";
      }
      onNotify({ kind: "success", text: "Файл добавлен в личную библиотеку." });
    } catch (uploadError: unknown) {
      onNotify({
        kind: "error",
        text: `Не удалось загрузить файл. ${getErrorMessage(uploadError)}`
      });
    } finally {
      setUploading(false);
    }
  };

  const insertAsset = async (item: PersonalAsset): Promise<void> => {
    if (insertingId !== null || deletingId !== null) {
      return;
    }

    setInsertingId(item.id);
    try {
      await powerPointService.insertPersonalAsset(item);
      onNotify({ kind: "success", text: insertionSuccessMessages[item.kind] });
    } catch (insertionError: unknown) {
      if (import.meta.env.DEV && !(insertionError instanceof PowerPointUnavailableError)) {
        console.error("Personal asset insertion failed", insertionError);
      }

      onNotify({
        kind: insertionError instanceof PowerPointUnavailableError ? "info" : "error",
        text:
          insertionError instanceof PowerPointUnavailableError
            ? insertionError.message
            : `Не удалось добавить материал в PowerPoint. ${getErrorMessage(insertionError)}`
      });
    } finally {
      setInsertingId(null);
    }
  };

  const deleteAsset = async (item: PersonalAsset): Promise<void> => {
    if (deletingId !== null || insertingId !== null) {
      return;
    }

    setDeletingId(item.id);
    try {
      await api.deletePersonalAsset(item.id);
      setItems((current) => current.filter((candidate) => candidate.id !== item.id));
      setConfirmingDeleteId(null);
      onNotify({ kind: "success", text: "Материал удалён из личной библиотеки." });
    } catch (deletionError: unknown) {
      if (import.meta.env.DEV) {
        console.error("Personal asset deletion failed", deletionError);
      }
      onNotify({
        kind: "error",
        text: `Не удалось удалить материал. ${getErrorMessage(deletionError)}`
      });
    } finally {
      setDeletingId(null);
    }
  };

  return (
    <section className="personal-library" aria-labelledby="personal-library-title">
      <header className="personal-library__header">
        <div>
          <span className="eyebrow">Личная библиотека</span>
          <h2 id="personal-library-title">{kindLabels[kind]}</h2>
          <p>Материалы видны только в вашей локальной библиотеке.</p>
        </div>
        <button
          className="icon-button"
          type="button"
          onClick={() => {
            setLoading(true);
            setError(null);
            setRefreshKey((value) => value + 1);
          }}
          disabled={loading}
          aria-label="Обновить личную библиотеку"
        >
          <svg viewBox="0 0 20 20" aria-hidden="true">
            <path d="M15.5 6.2A6 6 0 104 13.8M15.5 6.2V2.8M15.5 6.2h-3.4" />
          </svg>
        </button>
      </header>

      <form className="upload-panel" onSubmit={(event) => void submitUpload(event)}>
        <div className="upload-panel__intro">
          <strong>Добавить файл</strong>
          <span>{kindHints[kind]}</span>
        </div>
        <label htmlFor="personal-asset-title">
          <span>Название</span>
          <input
            id="personal-asset-title"
            value={title}
            onChange={(event) => setTitle(event.target.value)}
            maxLength={120}
            placeholder={`Название: ${kindLabels[kind].toLowerCase()}`}
            disabled={uploading}
          />
        </label>
        <label className="file-picker" htmlFor="personal-asset-file">
          <span>Файл</span>
          <input
            id="personal-asset-file"
            ref={fileInputRef}
            type="file"
            aria-label="Файл"
            accept={kindAccepts[kind]}
            onChange={(event) => setFile(event.target.files?.[0] ?? null)}
            disabled={uploading}
          />
          <span className="file-picker__surface">
            {file ? file.name : "Выбрать файл"}
          </span>
        </label>
        <button className="button button--primary" type="submit" disabled={uploading}>
          {uploading ? (
            <>
              <span className="spinner" aria-hidden="true" />
              Загружаем…
            </>
          ) : (
            "Добавить"
          )}
        </button>
      </form>

      {loading ? <p className="personal-library__message">Загружаем материалы…</p> : null}
      {!loading && error ? (
        <div className="state-panel state-panel--error" role="alert">
          <h3>Личная библиотека недоступна</h3>
          <p>{error}</p>
        </div>
      ) : null}
      {!loading && !error && visibleItems.length === 0 ? (
        <div className="state-panel">
          <h3>Здесь пока ничего нет</h3>
          <p>Загрузите первый файл с помощью формы выше.</p>
        </div>
      ) : null}
      {!loading && !error && visibleItems.length > 0 ? (
        <div className="personal-grid">
          {visibleItems.map((item) => (
            <article className="personal-card" key={item.id}>
              <div className="personal-card__preview">
                {item.mimeType.startsWith("image/") ? (
                  <img
                    src={api.getPersonalAssetFileUrl(item.id)}
                    alt={`${item.title} preview`}
                  />
                ) : (
                  <span aria-hidden="true">PPTX</span>
                )}
              </div>
              <div className="personal-card__body">
                <strong>{item.title}</strong>
                <span>{item.fileName}</span>
                <small>{formatFileSize(item.size)}</small>
                {confirmingDeleteId === item.id ? (
                  <span className="personal-card__confirmation">Удалить материал?</span>
                ) : null}
                <div
                  className={`personal-card__actions ${
                    confirmingDeleteId === item.id
                      ? "personal-card__actions--confirm"
                      : ""
                  }`}
                >
                  {confirmingDeleteId === item.id ? (
                    <>
                      <button
                        className="button button--secondary"
                        type="button"
                        onClick={() => setConfirmingDeleteId(null)}
                        disabled={deletingId !== null}
                      >
                        Отмена
                      </button>
                      <button
                        className="button button--danger"
                        type="button"
                        onClick={() => void deleteAsset(item)}
                        disabled={deletingId !== null}
                        aria-label={
                          deletingId === item.id
                            ? `Удаляем ${item.title}`
                            : `Подтвердить удаление ${item.title}`
                        }
                      >
                        {deletingId === item.id ? (
                          <>
                            <span className="spinner" aria-hidden="true" />
                            Удаляем…
                          </>
                        ) : (
                          "Удалить"
                        )}
                      </button>
                    </>
                  ) : (
                    <>
                      <button
                        className="button button--primary personal-card__action"
                        type="button"
                        onClick={() => void insertAsset(item)}
                        disabled={insertingId !== null || deletingId !== null}
                        aria-label={
                          insertingId === item.id
                            ? `Добавляем ${item.title} в PowerPoint`
                            : `Добавить ${item.title} в PowerPoint`
                        }
                      >
                        {insertingId === item.id ? (
                          <>
                            <span className="spinner" aria-hidden="true" />
                            Добавляем…
                          </>
                        ) : (
                          "Добавить"
                        )}
                      </button>
                      <button
                        className="button button--secondary personal-card__delete"
                        type="button"
                        onClick={() => setConfirmingDeleteId(item.id)}
                        disabled={insertingId !== null || deletingId !== null}
                        aria-label={`Удалить ${item.title} из личной библиотеки`}
                        title="Удалить"
                      >
                        <svg viewBox="0 0 20 20" aria-hidden="true">
                          <path d="M4.5 6h11M8 3.5h4M6.2 6l.7 10h6.2l.7-10M8.3 9v4.5M11.7 9v4.5" />
                        </svg>
                      </button>
                    </>
                  )}
                </div>
              </div>
            </article>
          ))}
        </div>
      ) : null}
    </section>
  );
}
