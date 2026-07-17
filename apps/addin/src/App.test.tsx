import { act, fireEvent, render, screen, waitFor, within } from "@testing-library/react";
import { describe, expect, it, vi } from "vitest";
import type { PersonalAsset, SlideListResponse } from "@slide-library/shared";
import { App } from "./App";
import {
  BROWSER_POWERPOINT_MESSAGE,
  BrowserPowerPointService
} from "./services/powerpoint";
import {
  catalogResponse,
  createApi,
  createAvailablePowerPointService,
  createDeferred,
  revenueSlide,
  strategySlide
} from "./test/fixtures";

const personalPhoto: PersonalAsset = {
  id: "11111111-1111-4111-8111-111111111111",
  title: "Командная фотография",
  kind: "photo",
  fileName: "team.png",
  mimeType: "image/png",
  size: 2048,
  createdAt: "2026-07-17T10:00:00.000Z"
};

describe("Slide Library application", () => {
  it("switches between presentations and the empty favorites section", async () => {
    render(<App api={createApi()} powerPointService={createAvailablePowerPointService()} />);

    expect(await screen.findByText(revenueSlide.title)).toBeInTheDocument();
    fireEvent.click(screen.getByRole("button", { name: "Избранное" }));

    expect(screen.getByRole("heading", { name: "Избранное пока пусто" })).toBeInTheDocument();
    expect(screen.queryByText(revenueSlide.title)).not.toBeInTheDocument();

    fireEvent.click(screen.getByRole("button", { name: "Открыть презентации" }));
    expect(screen.getByText(revenueSlide.title)).toBeInTheDocument();
  });

  it("adds and removes a slide from persistent favorites", async () => {
    render(<App api={createApi()} powerPointService={createAvailablePowerPointService()} />);

    await screen.findByText(revenueSlide.title);
    fireEvent.click(
      screen.getByRole("button", { name: `Add ${revenueSlide.title} to favorites` })
    );

    expect(globalThis.localStorage.getItem("slidebrary.favorite-slide-ids")).toContain(
      revenueSlide.id
    );

    fireEvent.click(screen.getByRole("button", { name: "Избранное" }));
    expect(screen.getByText(revenueSlide.title)).toBeInTheDocument();

    fireEvent.click(
      screen.getByRole("button", { name: `Remove ${revenueSlide.title} from favorites` })
    );
    expect(screen.getByRole("heading", { name: "Избранное пока пусто" })).toBeInTheDocument();
  });

  it("opens the personal library and uploads a photo", async () => {
    const api = createApi();
    render(<App api={api} powerPointService={createAvailablePowerPointService()} />);

    await screen.findByText(revenueSlide.title);
    fireEvent.click(screen.getByRole("button", { name: "Фотографии" }));

    expect(screen.getByRole("button", { name: "Личное" })).toHaveAttribute(
      "aria-current",
      "page"
    );
    expect(screen.getByRole("heading", { name: "Фотографии" })).toBeInTheDocument();

    fireEvent.change(screen.getByLabelText("Название"), {
      target: { value: "Командная фотография" }
    });
    const file = new File(
      [new Uint8Array([0x89, 0x50, 0x4e, 0x47])],
      "team.png",
      { type: "image/png" }
    );
    fireEvent.change(screen.getByLabelText("Файл"), {
      target: { files: [file] }
    });
    fireEvent.click(screen.getByRole("button", { name: "Добавить" }));

    await waitFor(() =>
      expect(api.uploadPersonalAsset).toHaveBeenCalledWith(
        "photo",
        "Командная фотография",
        file
      )
    );
    expect(await screen.findByText("Командная фотография")).toBeInTheDocument();
    expect(screen.getByText("Файл добавлен в личную библиотеку.")).toBeInTheDocument();
  });

  it("opens illustrations from the public library and uploads an illustration", async () => {
    const api = createApi();
    render(<App api={api} powerPointService={createAvailablePowerPointService()} />);

    await screen.findByText(revenueSlide.title);
    const illustrationsButton = screen.getByRole("button", { name: "Иллюстрации" });
    expect(illustrationsButton).toBeEnabled();
    fireEvent.click(illustrationsButton);

    expect(screen.getByRole("button", { name: "Личное" })).toHaveAttribute(
      "aria-current",
      "page"
    );
    expect(screen.getByRole("heading", { name: "Иллюстрации" })).toBeInTheDocument();

    fireEvent.change(screen.getByLabelText("Название"), {
      target: { value: "Иллюстрация продукта" }
    });
    const file = new File(
      ['<svg xmlns="http://www.w3.org/2000/svg"><path d="M0 0h1v1z"/></svg>'],
      "illustration.svg",
      { type: "image/svg+xml" }
    );
    fireEvent.change(screen.getByLabelText("Файл"), {
      target: { files: [file] }
    });
    fireEvent.click(screen.getByRole("button", { name: "Добавить" }));

    await waitFor(() =>
      expect(api.uploadPersonalAsset).toHaveBeenCalledWith(
        "illustration",
        "Иллюстрация продукта",
        file
      )
    );
    expect(await screen.findByText("Иллюстрация продукта")).toBeInTheDocument();
  });

  it("inserts a personal photo into PowerPoint and reports progress", async () => {
    const insertion = createDeferred<void>();
    const api = createApi();
    api.listPersonalAssets = vi.fn(async () => ({
      items: [personalPhoto],
      total: 1
    }));
    const insertPersonalAsset = vi.fn(() => insertion.promise);
    const powerPointService = createAvailablePowerPointService(
      async () => undefined,
      insertPersonalAsset
    );

    render(<App api={api} powerPointService={powerPointService} />);

    await screen.findByText(revenueSlide.title);
    fireEvent.click(screen.getByRole("button", { name: "Личное" }));
    fireEvent.click(screen.getByRole("button", { name: "Фотографии" }));
    const insertButton = await screen.findByRole("button", {
      name: `Добавить ${personalPhoto.title} в PowerPoint`
    });
    fireEvent.click(insertButton);

    expect(insertPersonalAsset).toHaveBeenCalledWith(personalPhoto);
    expect(
      screen.getByRole("button", {
        name: `Добавляем ${personalPhoto.title} в PowerPoint`
      })
    ).toBeDisabled();

    await act(async () => {
      insertion.resolve(undefined);
      await insertion.promise;
    });

    expect(
      await screen.findByText("Фотография добавлена на текущий слайд.")
    ).toBeInTheDocument();
    expect(
      screen.getByRole("button", {
        name: `Добавить ${personalPhoto.title} в PowerPoint`
      })
    ).toBeEnabled();
  });

  it("explains why personal assets cannot be inserted in browser mode", async () => {
    const api = createApi();
    api.listPersonalAssets = vi.fn(async () => ({
      items: [personalPhoto],
      total: 1
    }));

    render(<App api={api} powerPointService={new BrowserPowerPointService()} />);

    await screen.findByText(revenueSlide.title);
    fireEvent.click(screen.getByRole("button", { name: "Личное" }));
    fireEvent.click(screen.getByRole("button", { name: "Фотографии" }));
    fireEvent.click(
      await screen.findByRole("button", {
        name: `Добавить ${personalPhoto.title} в PowerPoint`
      })
    );

    const notification = await screen.findByRole("status");
    expect(within(notification).getByText(BROWSER_POWERPOINT_MESSAGE)).toBeInTheDocument();
  });

  it("requires confirmation and removes a personal asset from the library", async () => {
    const api = createApi();
    api.listPersonalAssets = vi.fn(async () => ({
      items: [personalPhoto],
      total: 1
    }));

    render(<App api={api} powerPointService={createAvailablePowerPointService()} />);

    await screen.findByText(revenueSlide.title);
    fireEvent.click(screen.getByRole("button", { name: "Личное" }));
    fireEvent.click(screen.getByRole("button", { name: "Фотографии" }));
    await screen.findByText(personalPhoto.title);

    const startDeletion = screen.getByRole("button", {
      name: `Удалить ${personalPhoto.title} из личной библиотеки`
    });
    fireEvent.click(startDeletion);
    expect(screen.getByText("Удалить материал?")).toBeInTheDocument();

    fireEvent.click(screen.getByRole("button", { name: "Отмена" }));
    expect(screen.queryByText("Удалить материал?")).not.toBeInTheDocument();

    fireEvent.click(
      screen.getByRole("button", {
        name: `Удалить ${personalPhoto.title} из личной библиотеки`
      })
    );
    fireEvent.click(
      screen.getByRole("button", {
        name: `Подтвердить удаление ${personalPhoto.title}`
      })
    );

    await waitFor(() =>
      expect(api.deletePersonalAsset).toHaveBeenCalledWith(personalPhoto.id)
    );
    expect(screen.queryByText(personalPhoto.title)).not.toBeInTheDocument();
    expect(
      await screen.findByText("Материал удалён из личной библиотеки.")
    ).toBeInTheDocument();
  });

  it("shows an accessible loading state until the catalog request completes", async () => {
    const catalog = createDeferred<SlideListResponse>();
    const api = createApi(() => catalog.promise);

    render(<App api={api} powerPointService={createAvailablePowerPointService()} />);

    expect(screen.getByRole("status", { name: "Loading slide library" })).toBeInTheDocument();
    expect(screen.getByRole("region", { name: "Slide catalog" })).toHaveAttribute(
      "aria-busy",
      "true"
    );

    await act(async () => {
      catalog.resolve(catalogResponse);
      await catalog.promise;
    });

    expect(await screen.findByText(revenueSlide.title)).toBeInTheDocument();
    expect(screen.queryByRole("status", { name: "Loading slide library" })).not.toBeInTheDocument();
  });

  it("renders catalog metadata and opens an enlarged, keyboard-dismissable details view", async () => {
    const api = createApi();

    render(<App api={api} powerPointService={createAvailablePowerPointService()} />);

    expect(await screen.findByText(revenueSlide.title)).toBeInTheDocument();
    const card = screen.getByRole("article");
    expect(within(card).getByText("Finance")).toBeInTheDocument();
    expect(within(card).getByText("Approved")).toBeInTheDocument();
    expect(screen.getByText("1", { selector: "strong" }).parentElement).toHaveTextContent(
      "1 slide"
    );
    expect(screen.getByRole("img", { name: `${revenueSlide.title} slide preview` })).toHaveAttribute(
      "src",
      "/api/slides/revenue-overview/preview"
    );

    fireEvent.click(
      screen.getByRole("button", { name: `View details for ${revenueSlide.title}` })
    );

    const dialog = screen.getByRole("dialog", { name: revenueSlide.title });
    expect(within(dialog).getByText("Version 2.4")).toBeInTheDocument();
    expect(within(dialog).getByText("Finance Operations")).toBeInTheDocument();
    expect(within(dialog).getByText("executive")).toBeInTheDocument();

    fireEvent.keyDown(document, { key: "Escape" });
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
  });

  it("debounces and trims search input before querying the API", async () => {
    const api = createApi();

    render(
      <App
        api={api}
        powerPointService={createAvailablePowerPointService()}
        searchDebounceMs={35}
      />
    );

    await screen.findByText(revenueSlide.title);
    expect(api.listSlides).toHaveBeenCalledTimes(1);

    fireEvent.change(screen.getByRole("searchbox", { name: "Search slides" }), {
      target: { value: "  revenue  " }
    });

    expect(api.listSlides).toHaveBeenCalledTimes(1);

    await waitFor(
      () => {
        expect(api.listSlides).toHaveBeenLastCalledWith(
          { query: "revenue", status: "approved" },
          expect.any(AbortSignal)
        );
      },
      { timeout: 500 }
    );
  });

  it("sends category and status filters and resets them to approved content", async () => {
    const api = createApi();

    render(
      <App
        api={api}
        powerPointService={createAvailablePowerPointService()}
        searchDebounceMs={0}
      />
    );

    await screen.findByText(revenueSlide.title);
    const category = screen.getByRole("combobox", { name: "Category" });
    const status = screen.getByRole("combobox", { name: "Status" });

    fireEvent.change(category, { target: { value: "Finance" } });
    await waitFor(() =>
      expect(api.listSlides).toHaveBeenLastCalledWith(
        { category: "Finance", status: "approved" },
        expect.any(AbortSignal)
      )
    );

    fireEvent.change(status, { target: { value: "draft" } });
    await waitFor(() =>
      expect(api.listSlides).toHaveBeenLastCalledWith(
        { category: "Finance", status: "draft" },
        expect.any(AbortSignal)
      )
    );

    fireEvent.click(screen.getByRole("button", { name: "Reset filters" }));
    expect(category).toHaveValue("");
    expect(status).toHaveValue("approved");
    await waitFor(() =>
      expect(api.listSlides).toHaveBeenLastCalledWith(
        { status: "approved" },
        expect.any(AbortSignal)
      )
    );
  });

  it("switches between grid and list views and sorts catalog items", async () => {
    const alphaSlide = { ...strategySlide, title: "Alpha strategy" };
    const api = createApi(async () => ({
      items: [alphaSlide, revenueSlide],
      total: 2,
      availableCategories: ["Finance", "Strategy"]
    }));

    render(<App api={api} powerPointService={createAvailablePowerPointService()} />);

    expect(await screen.findByText(revenueSlide.title)).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Плитка" })).toHaveAttribute(
      "aria-pressed",
      "true"
    );

    fireEvent.click(screen.getByRole("button", { name: "Список" }));
    expect(screen.getByRole("button", { name: "Список" })).toHaveAttribute(
      "aria-pressed",
      "true"
    );

    fireEvent.change(screen.getByRole("combobox", { name: "Sort slides" }), {
      target: { value: "title-asc" }
    });

    const cards = screen.getAllByRole("article");
    expect(cards).toHaveLength(2);
    expect(within(cards[0]!).getByText("Alpha strategy")).toBeInTheDocument();
  });

  it("shows no-results guidance when content exists but the current selection has no matches", async () => {
    const api = createApi(async () => ({
      items: [],
      total: 0,
      availableCategories: ["Finance"]
    }));

    render(<App api={api} powerPointService={createAvailablePowerPointService()} />);

    expect(await screen.findByRole("heading", { name: "No slides found" })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Clear search and filters" })).toBeInTheDocument();
  });

  it("distinguishes a genuinely empty library from filtered no-results", async () => {
    const api = createApi(async () => ({ items: [], total: 0, availableCategories: [] }));

    render(<App api={api} powerPointService={createAvailablePowerPointService()} />);

    expect(await screen.findByRole("heading", { name: "The library is empty" })).toBeInTheDocument();
    expect(
      within(screen.getByRole("region", { name: "Slide catalog" })).getByRole("button", {
        name: "Refresh library"
      })
    ).toBeInTheDocument();
  });

  it("recovers from an unavailable API when Retry succeeds", async () => {
    let attempts = 0;
    const api = createApi(async () => {
      attempts += 1;
      if (attempts === 1) {
        throw new Error("The catalog API is offline.");
      }
      return catalogResponse;
    });

    render(<App api={api} powerPointService={createAvailablePowerPointService()} />);

    const alert = await screen.findByRole("alert");
    expect(within(alert).getByText("The catalog API is offline.")).toBeInTheDocument();

    fireEvent.click(within(alert).getByRole("button", { name: "Retry" }));
    expect(await screen.findByText(revenueSlide.title)).toBeInTheDocument();
    expect(api.listSlides).toHaveBeenCalledTimes(2);
  });

  it("blocks all insert controls while insertion is pending and reports success", async () => {
    const insertion = createDeferred<void>();
    const insertSlide = vi.fn(() => insertion.promise);
    const api = createApi(async () => ({
      items: [revenueSlide, strategySlide],
      total: 2,
      availableCategories: ["Finance", "Strategy"]
    }));

    render(
      <App
        api={api}
        powerPointService={createAvailablePowerPointService(insertSlide)}
      />
    );

    const insertRevenue = await screen.findByRole("button", {
      name: `Insert ${revenueSlide.title}`
    });
    fireEvent.click(insertRevenue);

    expect(insertSlide).toHaveBeenCalledWith(revenueSlide.id);
    expect(
      screen.getByRole("button", { name: `Inserting ${revenueSlide.title}` })
    ).toBeDisabled();
    expect(screen.getByRole("button", { name: `Insert ${strategySlide.title}` })).toBeDisabled();

    await act(async () => {
      insertion.resolve(undefined);
      await insertion.promise;
    });

    expect(await screen.findByText("Slide inserted successfully")).toBeInTheDocument();
    expect(screen.getByRole("button", { name: `Insert ${revenueSlide.title}` })).toBeEnabled();
  });

  it("selects and inserts multiple slides in the visible catalog order", async () => {
    const api = createApi(async () => ({
      items: [strategySlide, revenueSlide],
      total: 2,
      availableCategories: ["Finance", "Strategy"]
    }));
    const powerPointService = createAvailablePowerPointService();

    render(<App api={api} powerPointService={powerPointService} />);

    await screen.findByText(revenueSlide.title);
    fireEvent.click(
      screen.getByRole("button", { name: `Select ${strategySlide.title}` })
    );
    fireEvent.click(
      screen.getByRole("button", { name: `Select ${revenueSlide.title}` })
    );

    fireEvent.click(screen.getByRole("button", { name: "Добавить выбранные (2)" }));

    await waitFor(() =>
      expect(powerPointService.insertSlides).toHaveBeenCalledWith([
        revenueSlide.id,
        strategySlide.id
      ])
    );
    expect(await screen.findByText("2 slides inserted successfully")).toBeInTheDocument();
    expect(screen.queryByLabelText("Выбранные слайды")).not.toBeInTheDocument();
  });

  it("reports insertion failures and restores the insert controls", async () => {
    vi.spyOn(console, "error").mockImplementation(() => undefined);
    const insertSlide = vi.fn(async () => {
      throw new Error("The PPTX file could not be downloaded.");
    });

    render(
      <App
        api={createApi()}
        powerPointService={createAvailablePowerPointService(insertSlide)}
      />
    );

    const insertButton = await screen.findByRole("button", {
      name: `Insert ${revenueSlide.title}`
    });
    fireEvent.click(insertButton);

    const notification = await screen.findByRole("alert");
    expect(notification).toHaveTextContent(
      "Could not insert slide. The PPTX file could not be downloaded."
    );
    expect(screen.getByRole("button", { name: `Insert ${revenueSlide.title}` })).toBeEnabled();
  });

  it("keeps browsing functional and gives a friendly insert explanation outside PowerPoint", async () => {
    const api = createApi();

    render(<App api={api} powerPointService={new BrowserPowerPointService()} />);

    await screen.findByText(revenueSlide.title);
    const integrationStatus = screen.getByLabelText("PowerPoint integration status");
    expect(within(integrationStatus).getByText("Catalog preview mode")).toBeInTheDocument();
    expect(within(integrationStatus).getByText(BROWSER_POWERPOINT_MESSAGE)).toBeInTheDocument();

    fireEvent.click(
      screen.getByRole("button", { name: `Insert ${revenueSlide.title}` })
    );

    const notification = await screen.findByRole("status");
    expect(within(notification).getByText(BROWSER_POWERPOINT_MESSAGE)).toBeInTheDocument();
  });
});
