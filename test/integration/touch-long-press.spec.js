const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, dispatchTouch, touchCenterOf } = require('../helpers');

// Long-press (~600ms hold) on row numbers, the # corner, or a column header
// opens the same context menus as right-click.

async function longPress(page, locator) {
  const handle = await locator.elementHandle();
  const { x, y } = await touchCenterOf(locator);
  await dispatchTouch(page, handle, 'touchstart', x, y);
  await page.waitForTimeout(750);
  await dispatchTouch(page, handle, 'touchend', x, y);
}

test.describe('Touch long-press context menus', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, 'sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('long-press on a row number opens the row context menu', async ({ page }) => {
    await longPress(page, page.locator('.subwindow td.row-num').first());
    const menu = page.locator('.context-menu');
    await expect(menu).toBeVisible();
    await expect(menu.locator('button', { hasText: 'Insert Row Below' })).toBeVisible();
    await expect(menu.locator('button', { hasText: 'Delete Row' })).toBeVisible();
  });

  test('long-press on a column header opens the column context menu', async ({ page }) => {
    await longPress(page, page.locator('.subwindow th[data-col-idx="0"] .col-name'));
    const menu = page.locator('.context-menu');
    await expect(menu).toBeVisible();
    await expect(menu.locator('button', { hasText: 'Rename Column…' })).toBeVisible();
    await expect(menu.locator('button', { hasText: 'Sort Ascending' })).toBeVisible();
    await expect(menu.locator('button', { hasText: 'Sort Descending' })).toBeVisible();
    await expect(menu.locator('button', { hasText: 'Insert Column Right' })).toBeVisible();
  });

  test('long-press on the # corner opens the corner context menu', async ({ page }) => {
    await longPress(page, page.locator('.subwindow th.row-num-header'));
    const menu = page.locator('.context-menu');
    await expect(menu).toBeVisible();
    await expect(menu.locator('button', { hasText: 'Rename Table…' })).toBeVisible();
  });

  test('moving the finger cancels the long-press', async ({ page }) => {
    const rowNum = page.locator('.subwindow td.row-num').first();
    const handle = await rowNum.elementHandle();
    const { x, y } = await touchCenterOf(rowNum);
    await dispatchTouch(page, handle, 'touchstart', x, y);
    await dispatchTouch(page, handle, 'touchmove', x, y + 30);
    await page.waitForTimeout(750);
    await dispatchTouch(page, handle, 'touchend', x, y + 30);
    await expect(page.locator('.context-menu')).toHaveCount(0);
  });

  test('a quick tap does not open a context menu', async ({ page }) => {
    const rowNum = page.locator('.subwindow td.row-num').first();
    const handle = await rowNum.elementHandle();
    const { x, y } = await touchCenterOf(rowNum);
    await dispatchTouch(page, handle, 'touchstart', x, y);
    await dispatchTouch(page, handle, 'touchend', x, y);
    await page.waitForTimeout(750);
    await expect(page.locator('.context-menu')).toHaveCount(0);
  });

  test('Sort Descending via long-press menu sorts the column', async ({ page }) => {
    await longPress(page, page.locator('.subwindow th[data-col-idx="0"] .col-name'));
    await page.locator('.context-menu button', { hasText: 'Sort Descending' }).click();
    const sortCols = await page.evaluate(() => app._test.windows[0].sortCols);
    expect(sortCols).toEqual([{ col: 'name', dir: 'desc' }]);
  });

  test('Add to Multi-sort and Remove from Sort appear contextually', async ({ page }) => {
    // Sort the first column, then open the second column's menu
    await longPress(page, page.locator('.subwindow th[data-col-idx="0"] .col-name'));
    await page.locator('.context-menu button', { hasText: 'Sort Ascending' }).click();
    await longPress(page, page.locator('.subwindow th[data-col-idx="1"] .col-name'));
    await page.locator('.context-menu button', { hasText: 'Add to Multi-sort' }).click();
    let sortCols = await page.evaluate(() => app._test.windows[0].sortCols);
    expect(sortCols).toEqual([{ col: 'name', dir: 'asc' }, { col: 'email', dir: 'asc' }]);
    // A sorted column's menu offers Remove from Sort instead
    await longPress(page, page.locator('.subwindow th[data-col-idx="1"] .col-name'));
    const menu = page.locator('.context-menu');
    await expect(menu.locator('button', { hasText: 'Remove from Sort' })).toBeVisible();
    await expect(menu.locator('button', { hasText: 'Add to Multi-sort' })).toHaveCount(0);
    await menu.locator('button', { hasText: 'Remove from Sort' }).click();
    sortCols = await page.evaluate(() => app._test.windows[0].sortCols);
    expect(sortCols).toEqual([{ col: 'name', dir: 'asc' }]);
  });

  test('Rename Column via right-click menu opens the inline rename input', async ({ page }) => {
    // Desktop parity: the new items are also reachable by mouse
    await page.locator('.subwindow th[data-col-idx="0"] .col-name').click({ button: 'right' });
    await page.locator('.context-menu button', { hasText: 'Rename Column…' }).click();
    await expect(page.locator('.subwindow th .inline-rename')).toBeVisible();
  });
});
