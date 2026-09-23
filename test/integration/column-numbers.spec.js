const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow } = require('../helpers');

// sample1.csv columns: name, email, member_since
test.describe('Column numbers', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  const statusRight = (page) => page.locator('.subwindow .win-statusbar .status-right');

  test('status bar shows the focused cell\'s column position', async ({ page }) => {
    await expect(statusRight(page)).toHaveText('3 columns');
    await page.locator('.subwindow tbody tr[data-display-idx="0"] td.data-cell').nth(1).click();
    await expect(statusRight(page)).toHaveText('Col 2 of 3');
    await page.keyboard.press('ArrowRight');
    await expect(statusRight(page)).toHaveText('Col 3 of 3');
    await page.keyboard.press('Escape');
    await expect(statusRight(page)).toHaveText('3 columns');
  });

  test('header column selection shows position alongside stats', async ({ page }) => {
    await page.locator('.subwindow thead th[data-col-idx="2"]').click();
    await expect(page.locator('.subwindow .status-col-pos')).toHaveText('Col 3 of 3');
    await expect(page.locator('.subwindow .status-stats')).toContainText('Count: 10');
  });

  test('multi-column selection shows a range', async ({ page }) => {
    await page.locator('.subwindow tbody tr[data-display-idx="0"] td.data-cell').first().click();
    await page.keyboard.press('Shift+ArrowRight');
    await page.keyboard.press('Shift+ArrowRight');
    await expect(page.locator('.subwindow .status-col-pos')).toHaveText('Cols 1–3 of 3');
  });

  test('position follows column reorder', async ({ page }) => {
    await page.locator('.subwindow thead th[data-col-idx="0"]').click();
    await expect(page.locator('.subwindow .status-col-pos')).toHaveText('Col 1 of 3');
    await page.keyboard.press('Control+ArrowRight');
    await expect(page.locator('.subwindow .status-col-pos')).toHaveText('Col 2 of 3');
  });

  test('View menu toggles numbered headers and persists the setting', async ({ page }) => {
    await expect(page.locator('.subwindow thead .col-num')).toHaveCount(0);
    await page.click('#menu-view .menu-label');
    await expect(page.locator('#btn-col-numbers')).toHaveText('Show Column Numbers');
    await page.click('#btn-col-numbers');
    await expect(page.locator('.subwindow thead .col-num')).toHaveText(['1', '2', '3']);
    await expect(page.locator('.subwindow table.data-table')).toHaveClass(/show-col-numbers/);
    expect(await page.evaluate(() => localStorage.getItem('csvsql_show_col_numbers'))).toBe('1');

    await page.click('#menu-view .menu-label');
    await expect(page.locator('#btn-col-numbers')).toHaveText('Hide Column Numbers');
    await page.click('#btn-col-numbers');
    await expect(page.locator('.subwindow thead .col-num')).toHaveCount(0);
    expect(await page.evaluate(() => localStorage.getItem('csvsql_show_col_numbers'))).toBe('0');
  });

  test('selected column number is emphasized in the header', async ({ page }) => {
    await page.click('#menu-view .menu-label');
    await page.click('#btn-col-numbers');
    await page.locator('.subwindow thead th[data-col-idx="1"]').click();
    const colNum = page.locator('.subwindow thead th[data-col-idx="1"] .col-num');
    await expect(colNum).toHaveCSS('font-weight', '600');
    await expect(page.locator('.subwindow thead th[data-col-idx="0"] .col-num')).toHaveCSS('font-weight', '400');
  });

  test('header numbers follow reorder and column insert', async ({ page }) => {
    await page.click('#menu-view .menu-label');
    await page.click('#btn-col-numbers');
    await page.locator('.subwindow thead th[data-col-idx="0"]').click();
    await page.keyboard.press('Control+ArrowRight');
    await expect(page.locator('.subwindow thead th[data-col-idx="1"] .col-name')).toHaveText('name');
    await expect(page.locator('.subwindow thead .col-num')).toHaveText(['1', '2', '3']);
    await page.locator('.subwindow thead th.row-num-header').click({ button: 'right' });
    await page.locator('.context-menu button', { hasText: 'Insert Column Right' }).click();
    await page.locator('.modal-input').fill('extra');
    await page.keyboard.press('Enter');
    await expect(page.locator('.subwindow thead .col-num')).toHaveText(['1', '2', '3', '4']);
    await expect(page.locator('.subwindow thead th[data-col-idx="0"] .col-name')).toHaveText('extra');
  });

  test('header grows by one short line with numbers on', async ({ page }) => {
    const theadH = () => page.locator('.subwindow thead').evaluate(el => el.offsetHeight);
    const before = await theadH();
    await page.click('#menu-view .menu-label');
    await page.click('#btn-col-numbers');
    await expect(page.locator('.subwindow thead .col-num')).toHaveCount(3);
    const after = await theadH();
    expect(after).toBeGreaterThan(before);
    expect(after - before).toBeLessThan(25);
    // Virtual scrolling still renders rows from the top after the header grows
    await expect(page.locator('.subwindow tbody tr[data-display-idx="0"]')).toBeVisible();
  });

  test('toggle applies to every open table window', async ({ page }) => {
    await uploadFile(page, '../test/customers.csv');
    await waitForWindow(page, 'customers');
    await page.click('#menu-view .menu-label');
    await page.click('#btn-col-numbers');
    for (const name of ['sample1', 'customers']) {
      const win = page.locator('.subwindow', { has: page.locator('.win-title', { hasText: name }) });
      await expect(win.locator('thead .col-num').first()).toHaveText('1');
    }
  });

  test('setting persists across reloads', async ({ page }) => {
    await page.click('#menu-view .menu-label');
    await page.click('#btn-col-numbers');
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
    await expect(page.locator('.subwindow thead .col-num')).toHaveText(['1', '2', '3']);
    await page.click('#menu-view .menu-label');
    await expect(page.locator('#btn-col-numbers')).toHaveText('Hide Column Numbers');
  });

  test('toggling keeps the focused cell and selection', async ({ page }) => {
    await page.locator('.subwindow tbody tr[data-display-idx="2"] td.data-cell').nth(1).click();
    await page.click('#menu-view .menu-label');
    await page.click('#btn-col-numbers');
    await expect(page.locator('.subwindow .status-col-pos')).toHaveText('Col 2 of 3');
    const focused = await page.evaluate(() => {
      const td = document.activeElement;
      return td && td.classList.contains('data-cell') ? td.closest('tr').dataset.displayIdx : null;
    });
    expect(focused).toBe('2');
  });

  test('View menu toggle is available with no windows open', async ({ page }) => {
    await openApp(page);
    await page.click('#menu-view .menu-label');
    await expect(page.locator('#btn-col-numbers')).toBeEnabled();
  });

  test('corner context menu no longer offers the toggle', async ({ page }) => {
    await page.locator('.subwindow thead th.row-num-header').click({ button: 'right' });
    await expect(page.locator('.context-menu')).toBeVisible();
    await expect(page.locator('.context-menu button', { hasText: 'Column Numbers' })).toHaveCount(0);
  });

  test('single-column range shows position before stats', async ({ page }) => {
    await page.locator('.subwindow tbody tr[data-display-idx="0"] td.data-cell').nth(2).click();
    await page.keyboard.press('Shift+ArrowDown');
    await expect(page.locator('.subwindow .status-col-pos')).toHaveText('Col 3 of 3');
    await expect(page.locator('.subwindow .status-stats')).toContainText('Count: 2');
  });
});

