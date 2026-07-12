const updatedDateFormatter = new Intl.DateTimeFormat("en-GB", {
  day: "2-digit",
  month: "short",
  year: "numeric",
  timeZone: "UTC"
});

export const formatUpdatedDate = (value: string): string => {
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? "Unknown date" : updatedDateFormatter.format(date);
};
