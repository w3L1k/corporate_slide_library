import { useEffect } from "react";

export interface ToastMessage {
  kind: "success" | "error" | "info";
  text: string;
}

interface ToastProps {
  message: ToastMessage;
  onDismiss(): void;
}

export function Toast({ message, onDismiss }: ToastProps) {
  useEffect(() => {
    const timeout = window.setTimeout(onDismiss, 5_000);
    return () => window.clearTimeout(timeout);
  }, [message, onDismiss]);

  return (
    <div
      className={`toast toast--${message.kind}`}
      role={message.kind === "error" ? "alert" : "status"}
      aria-live={message.kind === "error" ? "assertive" : "polite"}
    >
      <span className="toast__icon" aria-hidden="true">
        {message.kind === "success" ? "✓" : message.kind === "error" ? "!" : "i"}
      </span>
      <span>{message.text}</span>
      <button type="button" onClick={onDismiss} aria-label="Dismiss notification">
        <svg viewBox="0 0 16 16" aria-hidden="true">
          <path d="M4 4l8 8M12 4l-8 8" />
        </svg>
      </button>
    </div>
  );
}
