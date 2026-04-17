import { test, expect } from "@playwright/test";
import { readFileSync } from "fs";

test("generates and downloads a valid xlsx file", async ({ page }) => {
  await page.goto("/");

  const downloadPromise = page.waitForEvent("download");
  await page.click("#generate");
  const download = await downloadPromise;

  expect(download.suggestedFilename()).toBe("generated.xlsx");

  const filePath = await download.path();
  expect(filePath).toBeTruthy();
  const buf = readFileSync(filePath!);
  expect(buf.length).toBeGreaterThan(100);
  expect(buf[0]).toBe(0x50);
  expect(buf[1]).toBe(0x4b);

  await expect(page.locator("#status")).toHaveText("Done!");
});
