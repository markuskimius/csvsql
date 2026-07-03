const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow } = require('../helpers');

// sample1.csv columns: name, email, member_since (numeric). 3 data cells per row.
test.describe('Selection statistics', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  const statsEl = (page) => page.locator('.subwindow .win-statusbar .status-stats');

  test('no stats for a single-cell selection', async ({ page }) => {
    await page.locator('.subwindow table tbody td.data-cell').first().click();
    await expect(statsEl(page)).toHaveCount(0);
    // Row-count text is intact
    await expect(page.locator('.subwindow .win-statusbar .status-left')).toContainText('10 of 10 rows');
  });

  test('numeric selection shows Count, Sum, Avg, Min, Max', async ({ page }) => {
    // Select two cells in the numeric member_since column (3rd data cell of row 1)
    await page.locator('.subwindow table tbody tr[data-display-idx="0"] td.data-cell').nth(2).click();
    await page.keyboard.press('Shift+ArrowDown');
    const stats = statsEl(page);
    await expect(stats).toHaveCount(1);
    await expect(stats).toContainText('Count: 2');
    await expect(stats).toContainText('Sum:');
    await expect(stats).toContainText('Avg:');
    await expect(stats).toContainText('Min:');
    await expect(stats).toContainText('Max:');
    // Row-count text still present alongside the stats
    await expect(page.locator('.subwindow .win-statusbar .status-left')).toContainText('10 of 10 rows');
  });

  test('text-only selection shows Count without numeric stats', async ({ page }) => {
    await page.locator('.subwindow table tbody tr[data-display-idx="0"] td.data-cell').first().click();
    await page.keyboard.press('Shift+ArrowDown');
    const stats = statsEl(page);
    await expect(stats).toContainText('Count: 2');
    await expect(stats).not.toContainText('Sum:');
  });

  test('select-all computes stats across all cells', async ({ page }) => {
    await page.locator('.subwindow table tbody td.data-cell').first().click();
    await page.keyboard.press('Control+a');
    const stats = statsEl(page);
    await expect(stats).toContainText('Count: 30'); // 10 rows × 3 columns, all non-empty
    await expect(stats).toContainText('Sum:');      // member_since column is numeric
  });

  test('clearing the selection removes the stats', async ({ page }) => {
    await page.locator('.subwindow table tbody td.data-cell').first().click();
    await page.keyboard.press('Shift+ArrowDown');
    await expect(statsEl(page)).toHaveCount(1);
    await page.keyboard.press('Escape'); // select mode Esc clears the selection
    await expect(statsEl(page)).toHaveCount(0);
  });

  test('stats update after cut', async ({ page, context }) => {
    await context.grantPermissions(['clipboard-read', 'clipboard-write']);
    await page.locator('.subwindow table tbody tr[data-display-idx="0"] td.data-cell').nth(2).click();
    await page.keyboard.press('Shift+ArrowDown');
    await expect(statsEl(page)).toContainText('Count: 2');
    await page.keyboard.press('Control+x');
    await page.waitForTimeout(200);
    // Both cells are now empty — Count drops to 0 and numeric stats disappear
    await expect(statsEl(page)).toContainText('Count: 0');
    await expect(statsEl(page)).not.toContainText('Sum:');
  });
});
