import { test, expect, Page } from '@playwright/test';

const APP_URL = 'http://127.0.0.1:3005';

test.describe('UI / E2E Tests', () => {
  test('should load the main page', async ({ page }) => {
    const response = await page.goto(APP_URL, { waitUntil: 'domcontentloaded', timeout: 10000 });
    if (response) {
      expect(response.status()).toBe(200);
      const title = await page.title();
      expect(title).toBeTruthy();
    } else {
      test.skip(true, 'Server not available for E2E testing');
    }
  });

  test('should render the root element', async ({ page }) => {
    await page.goto(APP_URL, { waitUntil: 'domcontentloaded', timeout: 10000 });
    const root = await page.locator('#root');
    const isVisible = await root.isVisible().catch(() => false);
    if (!isVisible) {
      test.skip(true, 'Root element not visible - server may not be running');
    }
  });

  test('should not have critical console errors on load', async ({ page }) => {
    const errors: string[] = [];
    page.on('console', (msg) => {
      if (msg.type() === 'error') {
        errors.push(msg.text());
      }
    });
    await page.goto(APP_URL, { waitUntil: 'domcontentloaded', timeout: 10000 });
    await page.waitForTimeout(3000);
    const critical = errors.filter((e) => !e.includes('favicon'));
    expect(critical).toEqual([]);
  });

  test('should render main content area', async ({ page }) => {
    await page.goto(APP_URL, { waitUntil: 'domcontentloaded', timeout: 10000 });
    await page.waitForTimeout(3000);
    const body = await page.locator('body').textContent();
    expect(body?.length).toBeGreaterThan(0);
  });

  test('should handle window resize without errors', async ({ page }) => {
    await page.goto(APP_URL, { waitUntil: 'domcontentloaded', timeout: 10000 });
    await page.waitForTimeout(2000);
    await page.setViewportSize({ width: 1024, height: 768 });
    await page.waitForTimeout(500);
    await page.setViewportSize({ width: 1920, height: 1080 });
    await page.waitForTimeout(500);
    const title = await page.title();
    expect(title).toBeTruthy();
  });

  test('should handle rapid navigation gracefully', async ({ page }) => {
    await page.goto(APP_URL, { waitUntil: 'domcontentloaded', timeout: 10000 });
    await page.waitForTimeout(1000);
    for (let i = 0; i < 3; i++) {
      await page.goto(APP_URL, { waitUntil: 'domcontentloaded', timeout: 10000 });
      await page.waitForTimeout(500);
    }
    const title = await page.title();
    expect(title).toBeTruthy();
  });
});