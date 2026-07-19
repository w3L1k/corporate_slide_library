const CATEGORY_LABELS: Readonly<Record<string, string>> = {
  Charts: "Графики",
  Company: "Компания",
  Finance: "Финансы",
  Marketing: "Маркетинг",
  Product: "Продукт",
  Strategy: "Стратегия"
};

export const formatCategory = (category: string): string =>
  CATEGORY_LABELS[category] ?? category;
