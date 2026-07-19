import { useState } from "react";

interface PreviewImageProps {
  src: string;
  title: string;
  eager?: boolean;
  className?: string;
}

function PreviewPlaceholder({ title }: Pick<PreviewImageProps, "title">) {
  return (
    <div
      className="preview-placeholder"
      role="img"
      aria-label={`Предпросмотр недоступен: ${title}`}
    >
      <svg viewBox="0 0 48 48" aria-hidden="true">
        <rect x="7" y="10" width="34" height="25" rx="3" />
        <path d="M14 27l6-6 5 5 4-4 7 7" />
        <circle cx="31" cy="18" r="2.5" />
        <path d="M19 40h10" />
      </svg>
      <span>Предпросмотр недоступен</span>
    </div>
  );
}

export function PreviewImage({ src, title, eager = false, className = "" }: PreviewImageProps) {
  const [failed, setFailed] = useState(false);

  if (failed) {
    return <PreviewPlaceholder title={title} />;
  }

  return (
    <img
      className={className}
      src={src}
      alt={`Предпросмотр слайда: ${title}`}
      loading={eager ? "eager" : "lazy"}
      decoding="async"
      onError={() => setFailed(true)}
    />
  );
}
