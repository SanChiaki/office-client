import { act, cleanup, render, screen, waitFor } from "@testing-library/react";
import { afterEach, expect, test, vi } from "vitest";
import App from "../../src/App";

afterEach(() => {
  cleanup();
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

test("updates the selection badge when Office reports a new selection", async () => {
  let registeredHandler: ((eventArgs: unknown) => Promise<void> | void) | null = null;
  const addHandlerAsync = vi.fn((eventType: string, handler: typeof registeredHandler, callback?: (result: { status: string }) => void) => {
    registeredHandler = handler;
    callback?.({ status: "succeeded" });
  });
  const removeHandlerAsync = vi.fn((eventType: string, payload: { handler: typeof registeredHandler }, callback?: (result: { status: string }) => void) => {
    callback?.({ status: "succeeded" });
  });
  const load = vi.fn();
  const sync = vi.fn(async () => {});
  const getSelectedRange = vi.fn(() => ({
    address: "D4:E6",
    rowCount: 3,
    columnCount: 2,
    load,
    worksheet: {
      load: vi.fn(),
      name: "Sheet7",
    },
  }));

  vi.stubGlobal("Office", {
    context: {
      document: {
        addHandlerAsync,
        removeHandlerAsync,
      },
    },
  });
  vi.stubGlobal("Excel", {
    run: vi.fn(async (callback: (context: { workbook: { getSelectedRange: typeof getSelectedRange }; sync: typeof sync }) => Promise<void>) =>
      callback({
        workbook: { getSelectedRange },
        sync,
      }),
    ),
  });

  render(<App />);

  expect(screen.getByText("当前选区：未选择")).toBeInTheDocument();
  expect(registeredHandler).toEqual(expect.any(Function));

  await act(async () => {
    await registeredHandler?.({ document: {} });
  });

  await waitFor(() => {
    const badge = document.querySelector(".selection-badge");
    expect(badge?.textContent).toContain("Sheet7!D4:E6");
    expect(badge?.textContent).toContain("3");
    expect(badge?.textContent).toContain("2");
  });

  expect(removeHandlerAsync).not.toHaveBeenCalled();
});

test("hydrates the selection badge on mount when Office is ready", async () => {
  const addHandlerAsync = vi.fn((eventType: string, handler: (eventArgs: unknown) => Promise<void> | void, callback?: (result: { status: string }) => void) => {
    callback?.({ status: "succeeded" });
  });
  const removeHandlerAsync = vi.fn();
  const load = vi.fn();
  const worksheetLoad = vi.fn();
  const sync = vi.fn(async () => {});
  const getSelectedRange = vi.fn(() => ({
    address: "A1:D4",
    rowCount: 4,
    columnCount: 4,
    load,
    worksheet: {
      load: worksheetLoad,
      name: "Sheet1",
    },
  }));

  vi.stubGlobal("Office", {
    onReady: () => Promise.resolve(),
    context: {
      document: {
        addHandlerAsync,
        removeHandlerAsync,
      },
    },
  });
  vi.stubGlobal("Excel", {
    run: vi.fn(async (callback: (context: { workbook: { getSelectedRange: typeof getSelectedRange }; sync: typeof sync }) => Promise<void>) =>
      callback({
        workbook: { getSelectedRange },
        sync,
      }),
    ),
  });

  render(<App />);

  await waitFor(() => {
    const badges = document.querySelectorAll(".selection-badge");
    const latestBadgeText = badges[badges.length - 1]?.textContent ?? "";
    expect(latestBadgeText).toMatch(/Sheet1!A1:D4.*4.*4/);
  });

  expect(addHandlerAsync).toHaveBeenCalledTimes(1);
  expect(load).toHaveBeenCalledWith(["address", "rowCount", "columnCount"]);
  expect(worksheetLoad).toHaveBeenCalledWith("name");
  expect(sync).toHaveBeenCalledTimes(1);
});
