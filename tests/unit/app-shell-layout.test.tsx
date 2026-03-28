import { render } from "@testing-library/react";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { expect, test } from "vitest";
import App from "../../src/App";

test("task pane shell keeps the message thread scrollable and adds a narrow-width fallback", () => {
  render(<App />);

  const cssText = readFileSync(resolve(process.cwd(), "src/styles.css"), "utf8");

  expect(cssText).toContain("@media (max-width:");
  expect(cssText).toContain(".message-thread");
  expect(cssText).toContain("overflow: auto");
  expect(cssText).toContain("minmax(0, 1fr)");
});
