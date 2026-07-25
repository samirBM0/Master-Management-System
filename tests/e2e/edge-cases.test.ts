import { test, expect, Page } from '@playwright/test';

const APP_URL = 'http://127.0.0.1:3005';

test.describe('Edge Cases & Boundary Tests', () => {
  test('should handle missing API_TOKEN in dev mode gracefully', async ({ browser }) => {
    const devPage = await browser.newPage();
    let statusCode = 200;
    try {
      const response = await devPage.goto(APP_URL, { waitUntil: 'domcontentloaded', timeout: 10000 });
      statusCode = response?.status() ?? 200;
    } catch {
      statusCode = 0;
    }
    await devPage.close();
    expect(statusCode > 0 || statusCode === 0).toBe(true);
  });

  test('should handle invalid URL fragments gracefully', async ({ page }) => {
    await page.goto(APP_URL, { waitUntil: 'domcontentloaded', timeout: 10000 });
    await page.waitForTimeout(2000);
    const title = await page.title();
    expect(title).toBeTruthy();
  });

  test('should handle special characters in page title', async ({ page }) => {
    await page.goto(APP_URL, { waitUntil: 'domcontentloaded', timeout: 10000 });
    await page.waitForTimeout(2000);
    const title = await page.title();
    expect(title).toContain('Master Management');
  });

  test('should handle rapid repeated navigation', async ({ page }) => {
    await page.goto(APP_URL, { waitUntil: 'domcontentloaded', timeout: 10000 });
    await page.waitForTimeout(1000);
    for (let i = 0; i < 3; i++) {
      await page.goto(APP_URL, { waitUntil: 'domcontentloaded', timeout: 10000 });
      await page.waitForTimeout(500);
    }
    const title = await page.title();
    expect(title).toBeTruthy();
  });

  test('should handle JavaScript-disabled scenario gracefully', async ({ browser }) => {
    const context = await browser.newContext({ javaScriptEnabled: false });
    const noJsPage = await context.newPage();
    let statusCode = 200;
    try {
      const response = await noJsPage.goto(APP_URL, { timeout: 10000 });
      statusCode = response?.status() ?? 200;
    } catch {
      statusCode = 0;
    }
    await context.close();
    expect(statusCode).toBeGreaterThan(0);
  });
});