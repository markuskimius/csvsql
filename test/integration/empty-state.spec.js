const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow } = require('../helpers');

test.describe('Empty state', () => {
  test('visible on startup, hidden after opening a file, reappears after closing', async ({ page }) => {
    await openApp(page);
    await expect(page.locator('#empty-state')).toBeVisible();

    await uploadFile(page, 'sample1.csv');
    await waitForWindow(page, 'sample1');
    await expect(page.locator('#empty-state')).toBeHidden();

    await page.locator('.subwindow .btn-close').first().click();
    await expect(page.locator('#empty-state')).toBeVisible();
  });

  test('stays hidden while a window is merely minimized', async ({ page }) => {
    await openApp(page);
    await uploadFile(page, 'sample1.csv');
    await waitForWindow(page, 'sample1');
    await page.locator('.subwindow .btn-min').first().click();
    await expect(page.locator('.subwindow')).toHaveClass(/minimized/);
    await expect(page.locator('#empty-state')).toBeHidden();
  });

  test('Try Example Data loads sample tables and hides the empty state', async ({ page }) => {
    await openApp(page);
    await page.locator('#empty-state button', { hasText: 'Try Example Data' }).click();
    await waitForWindow(page, 'customers');
    await waitForWindow(page, 'orders');
    await waitForWindow(page, 'products');
    await expect(page.locator('#empty-state')).toBeHidden();

    const tableNames = await page.evaluate(() => Object.keys(app._test.tables).sort());
    expect(tableNames).toEqual(['customers', 'orders', 'products']);

    // Example tables are queryable and joinable in SQLite
    const rowCount = await page.evaluate(() => {
      const res = app._test.db.exec('SELECT COUNT(*) FROM orders o JOIN customers c ON o.customer_id = c.id');
      return res[0].values[0][0];
    });
    expect(rowCount).toBe(7);
  });
});
