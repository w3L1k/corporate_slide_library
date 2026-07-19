import type { SlideStatus } from "@slide-library/shared";

const STATUS_LABELS: Record<SlideStatus, string> = {
  approved: "Одобрено",
  draft: "Черновик",
  deprecated: "Устарело"
};

interface StatusBadgeProps {
  status: SlideStatus;
}

export function StatusBadge({ status }: StatusBadgeProps) {
  return <span className={`status-badge status-badge--${status}`}>{STATUS_LABELS[status]}</span>;
}
