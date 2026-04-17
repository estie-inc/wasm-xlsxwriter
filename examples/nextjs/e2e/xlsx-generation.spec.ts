import { test, expect } from "@playwright/test";

test.describe("App Router", () => {
  test("client-side generation", async ({ page }) => {
    await page.goto("/app-client");
    await page.click('[data-testid="generate"]');
    const status = page.locator('[data-testid="status"]');
    await expect(status).toContainText("Generated", { timeout: 10000 });
    await expect(status).toContainText("bytes");
  });

  test("route handler (server-side)", async ({ request }) => {
    const response = await request.get("/api/generate");
    expect(response.status()).toBe(200);
    expect(response.headers()["content-type"]).toContain(
      "spreadsheetml.sheet"
    );
    const body = await response.body();
    expect(body.length).toBeGreaterThan(0);
    // PK zip header (xlsx is a zip archive)
    expect(body[0]).toBe(0x50);
    expect(body[1]).toBe(0x4b);
  });
});

test.describe("Pages Router", () => {
  test("client-side generation", async ({ page }) => {
    await page.goto("/pages-client");
    await page.click('[data-testid="generate"]');
    const status = page.locator('[data-testid="status"]');
    await expect(status).toContainText("Generated", { timeout: 10000 });
    await expect(status).toContainText("bytes");
  });

  test("API route (server-side)", async ({ request }) => {
    const response = await request.get("/api/pages-generate");
    expect(response.status()).toBe(200);
    expect(response.headers()["content-type"]).toContain(
      "spreadsheetml.sheet"
    );
    const body = await response.body();
    expect(body.length).toBeGreaterThan(0);
    expect(body[0]).toBe(0x50);
    expect(body[1]).toBe(0x4b);
  });
});
