const SKELETON_CARD_COUNT = 3;

export function SkeletonCatalog() {
  return (
    <div className="skeleton-catalog" role="status" aria-label="Загрузка библиотеки слайдов">
      <span className="sr-only">Загружаем библиотеку слайдов…</span>
      {Array.from({ length: SKELETON_CARD_COUNT }, (_, index) => (
        <div className="skeleton-card" aria-hidden="true" key={index}>
          <div className="skeleton skeleton-card__preview" />
          <div className="skeleton-card__body">
            <div className="skeleton skeleton-card__title" />
            <div className="skeleton skeleton-card__meta" />
            <div className="skeleton skeleton-card__button" />
          </div>
        </div>
      ))}
    </div>
  );
}
