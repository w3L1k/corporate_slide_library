const updatedDateFormatter = new Intl.DateTimeFormat("ru-RU", {
  day: "2-digit",
  month: "short",
  year: "numeric",
  timeZone: "UTC"
});

export const formatUpdatedDate = (value: string): string => {
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? "Дата неизвестна" : updatedDateFormatter.format(date);
};
