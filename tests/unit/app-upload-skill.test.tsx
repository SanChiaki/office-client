import { afterEach, beforeEach, expect, test, vi } from "vitest";
import { cleanup, fireEvent, render, screen, waitFor, within } from "@testing-library/react";
import App from "../../src/App";

beforeEach(() => {
  window.localStorage.clear();
  window.localStorage.setItem(
    "oa:settings",
    JSON.stringify({
      apiKey: "sk-demo",
      baseUrl: "https://api.example.com",
      model: "gpt-4.1-mini",
    }),
  );
});

afterEach(() => {
  cleanup();
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

test("builds an upload preview and submits it after confirmation", async () => {
  const fetchMock = vi.fn(async () => ({
    ok: true,
    json: vi.fn(async () => ({
      saved: 2,
      project: "项目A",
    })),
  }));
  const excelAdapter = {
    readSelectionTable: vi.fn(async () => ({
      headers: ["Name", "Owner"],
      rows: [
        ["项目A", "张三"],
        ["项目B", "李四"],
      ],
    })),
    run: vi.fn(async (action: { type: string }) => action.type),
  };

  vi.stubGlobal("fetch", fetchMock);

  render(<App excelAdapterFactory={() => excelAdapter} />);

  const input = screen.getByRole("textbox");
  const sendButton = document.querySelector(".composer-send");
  if (!sendButton) {
    throw new Error("Send button was not rendered");
  }

  fireEvent.change(input, { target: { value: "把选中数据上传到项目A" } });
  fireEvent.click(sendButton);

  await waitFor(() => {
    expect(screen.getByText(/Name/)).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "确认" })).toBeInTheDocument();
  });

  fireEvent.click(screen.getByRole("button", { name: "确认" }));

  await waitFor(() => {
    expect(fetchMock).toHaveBeenCalledWith("https://api.example.com/upload_data_api", {
      method: "POST",
      headers: {
        Authorization: "Bearer sk-demo",
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        project: "项目A",
        columns: ["Name", "Owner"],
        rowCount: 2,
        previewRows: [
          ["项目A", "张三"],
          ["项目B", "李四"],
        ],
      }),
    });
  });

  const thread = screen.getByRole("region", { name: "消息线程" });
  expect(within(thread).getByText("把选中数据上传到项目A")).toBeInTheDocument();
  expect(within(thread).getByText(/上传完成/)).toBeInTheDocument();
  expect(excelAdapter.readSelectionTable).toHaveBeenCalledTimes(1);
  expect(excelAdapter.run).not.toHaveBeenCalled();
});

test("asks for an API key before starting the upload skill", async () => {
  window.localStorage.setItem(
    "oa:settings",
    JSON.stringify({
      apiKey: "",
      baseUrl: "https://api.example.com",
      model: "gpt-4.1-mini",
    }),
  );

  const excelAdapter = {
    readSelectionTable: vi.fn(async () => ({
      headers: ["Name", "Owner"],
      rows: [["项目A", "张三"]],
    })),
    run: vi.fn(async (action: { type: string }) => action.type),
  };

  render(<App excelAdapterFactory={() => excelAdapter} />);

  fireEvent.change(screen.getByRole("textbox"), { target: { value: "把选中数据上传到项目A" } });
  fireEvent.click(document.querySelector(".composer-send")!);

  const thread = screen.getByRole("region", { name: "消息线程" });
  expect(await within(thread).findByText("请先在设置中填写 API Key。")).toBeInTheDocument();
  expect(excelAdapter.readSelectionTable).not.toHaveBeenCalled();
});

test("blocks upload skill execution for non-contiguous selections", async () => {
  const addHandlerAsync = vi.fn((eventType: string, handler: (eventArgs: unknown) => Promise<void> | void, callback?: (result: { status: string }) => void) => {
    callback?.({ status: "succeeded" });
  });
  const removeHandlerAsync = vi.fn();
  const load = vi.fn();
  const worksheetLoad = vi.fn();
  const sync = vi.fn(async () => {});
  const getSelectedRange = vi.fn(() => ({
    address: "A1:B2,D1:E2",
    rowCount: 2,
    columnCount: 4,
    load,
    worksheet: {
      load: worksheetLoad,
      name: "Sheet1",
    },
  }));
  const excelAdapter = {
    readSelectionTable: vi.fn(async () => ({
      headers: ["Name", "Owner"],
      rows: [["项目A", "张三"]],
    })),
    run: vi.fn(async (action: { type: string }) => action.type),
  };

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

  render(<App excelAdapterFactory={() => excelAdapter} />);

  await waitFor(() => {
    expect(screen.getByText(/Sheet1!A1:B2,D1:E2/)).toBeInTheDocument();
  });

  fireEvent.change(screen.getByRole("textbox"), { target: { value: "把选中数据上传到项目A" } });
  fireEvent.click(document.querySelector(".composer-send")!);

  const thread = screen.getByRole("region", { name: "消息线程" });
  expect(await within(thread).findByText("首版仅支持连续选区，请重新选择单个连续区域。")).toBeInTheDocument();
  expect(excelAdapter.readSelectionTable).not.toHaveBeenCalled();
});
