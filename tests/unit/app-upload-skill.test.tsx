import { afterEach, beforeEach, expect, test, vi } from "vitest";
import { fireEvent, render, screen, waitFor, within } from "@testing-library/react";
import App from "../../src/App";

beforeEach(() => {
  window.localStorage.clear();
  window.localStorage.setItem(
    "oa:settings",
    JSON.stringify({
      apiKey: "sk-demo",
      model: "gpt-4.1-mini",
    }),
  );
});

afterEach(() => {
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
