const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, executeSQL } = require('../helpers');

async function getWinState(page, tableName) {
  return page.evaluate((name) => {
    const w = app._test.windows.find(w => w.tableName === name);
    if (!w) return null;
    return {
      selectedCol: w.selectedCol,
      sortCols: w.sortCols,
      selectedCells: [...w.selectedCells].sort(),
      anchorCell: w.anchorCell,
      copyWithHeader: w._copyWithHeader,
    };
  }, tableName);
}

test.describe('Column selection', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  // --- Column Selection via Header Click ---

  test('header click selects all cells in that column', async ({ page }) => {
    const header = page.locator('.subwindow table thead th').nth(1);
    await header.click();
    await page.waitForTimeout(100);

    const state = await getWinState(page, 'sample1');
    expect(state.selectedCells.length).toBe(10);
    expect(state.selectedCells.every(c => c.endsWith(':name'))).toBe(true);
    expect(state.copyWithHeader).toBe(true);
  });

  test('Escape clears a header-click column selection', async ({ page }) => {
    // A header click selects the column without moving focus to any cell, so
    // Escape is handled by the document-level fallback, not the grid handler.
    const header = page.locator('.subwindow table thead th').nth(1);
    await header.click();
    await page.waitForTimeout(100);
    let state = await getWinState(page, 'sample1');
    expect(state.selectedCells.length).toBe(10);

    await page.keyboard.press('Escape');
    await page.waitForTimeout(100);
    state = await getWinState(page, 'sample1');
    expect(state.selectedCells.length).toBe(0);
    expect(state.anchorCell).toBe(null);
  });

  test('header click sets anchorCell to first display row of clicked column', async ({ page }) => {
    const header = page.locator('.subwindow table thead th').nth(2); // email
    await header.click();
    await page.waitForTimeout(100);

    const state = await getWinState(page, 'sample1');
    expect(state.anchorCell).toEqual({ rownum: 1, col: 'email' });
  });

  test('header click sets selectedCol', async ({ page }) => {
    const header = page.locator('.subwindow table thead th').nth(1);
    await header.click();
    await page.waitForTimeout(100);

    const state = await getWinState(page, 'sample1');
    expect(state.selectedCol).toBe('name');
  });

  test('header click does not sort', async ({ page }) => {
    const header = page.locator('.subwindow table thead th').nth(1);
    await header.click();
    await page.waitForTimeout(100);

    const state = await getWinState(page, 'sample1');
    expect(state.sortCols).toEqual([]);
  });

  test('header click adds col-highlight class to the th', async ({ page }) => {
    const header = page.locator('.subwindow table thead th').nth(1);
    await header.click();
    await page.waitForTimeout(100);

    await expect(header).toHaveClass(/col-highlight/);
  });

  test('clicking a different header replaces the selection', async ({ page }) => {
    const nameHeader = page.locator('.subwindow table thead th').nth(1);
    const emailHeader = page.locator('.subwindow table thead th').nth(2);

    await nameHeader.click();
    await page.waitForTimeout(100);

    await emailHeader.click();
    await page.waitForTimeout(100);

    const state = await getWinState(page, 'sample1');
    expect(state.selectedCol).toBe('email');
    expect(state.selectedCells.length).toBe(10);
    expect(state.selectedCells.every(c => c.endsWith(':email'))).toBe(true);
    // name header should no longer be highlighted
    await expect(nameHeader).not.toHaveClass(/col-highlight/);
    await expect(emailHeader).toHaveClass(/col-highlight/);
  });

  test('Shift+click extends column selection to a range', async ({ page }) => {
    const nameHeader = page.locator('.subwindow table thead th').nth(1);
    const memberHeader = page.locator('.subwindow table thead th').nth(3);

    await nameHeader.click();
    await page.waitForTimeout(100);

    await memberHeader.click({ modifiers: ['Shift'] });
    await page.waitForTimeout(100);

    const state = await getWinState(page, 'sample1');
    // All 3 columns x 10 rows = 30 cells
    expect(state.selectedCells.length).toBe(30);
    const cols = new Set(state.selectedCells.map(c => c.split(':')[1]));
    expect([...cols].sort()).toEqual(['email', 'member_since', 'name']);
  });

  test('Ctrl+click on header triggers rename, not select', async ({ page }) => {
    const header = page.locator('.subwindow table thead th').nth(1);
    const mod = process.platform === 'darwin' ? 'Meta' : 'Control';
    await header.click({ modifiers: [mod] });
    await page.waitForTimeout(200);

    const renameInput = header.locator('input.inline-rename');
    await expect(renameInput).toBeVisible();
  });

  // --- Column Selection via Context Menu ---

  test('right-click on header selects that column', async ({ page }) => {
    const header = page.locator('.subwindow table thead th').nth(2); // email
    await header.click({ button: 'right' });
    await page.waitForTimeout(200);

    const state = await getWinState(page, 'sample1');
    expect(state.selectedCells.length).toBe(10);
    expect(state.selectedCells.every(c => c.endsWith(':email'))).toBe(true);
  });

  test('right-click on header sets selectedCol', async ({ page }) => {
    const header = page.locator('.subwindow table thead th').nth(2);
    await header.click({ button: 'right' });
    await page.waitForTimeout(200);

    const state = await getWinState(page, 'sample1');
    expect(state.selectedCol).toBe('email');
  });

  // --- Sort Badge Behavior ---

  test('sort badge click sorts ascending', async ({ page }) => {
    const header = page.locator('.subwindow table thead th').nth(1);
    const badge = header.locator('.col-badge-sort');
    await badge.click();
    await page.waitForTimeout(200);

    const state = await getWinState(page, 'sample1');
    expect(state.sortCols).toEqual([{ col: 'name', dir: 'asc' }]);
  });

  test('sort badge click toggles asc then desc then unsorted', async ({ page }) => {
    const header = page.locator('.subwindow table thead th').nth(1);
    const badge = header.locator('.col-badge-sort');

    await badge.click();
    await page.waitForTimeout(200);
    let state = await getWinState(page, 'sample1');
    expect(state.sortCols[0].dir).toBe('asc');

    await badge.click();
    await page.waitForTimeout(200);
    state = await getWinState(page, 'sample1');
    expect(state.sortCols[0].dir).toBe('desc');

    await badge.click();
    await page.waitForTimeout(200);
    state = await getWinState(page, 'sample1');
    expect(state.sortCols).toEqual([]);
  });

  test('sort badge click does not set selectedCol', async ({ page }) => {
    const header = page.locator('.subwindow table thead th').nth(1);
    const badge = header.locator('.col-badge-sort');
    await badge.click();
    await page.waitForTimeout(200);

    const state = await getWinState(page, 'sample1');
    expect(state.selectedCol).toBeNull();
  });

  test('sort badge Shift+click adds to multi-sort', async ({ page }) => {
    const nameBadge = page.locator('.subwindow table thead th').nth(1).locator('.col-badge-sort');
    const emailBadge = page.locator('.subwindow table thead th').nth(2).locator('.col-badge-sort');

    await nameBadge.click();
    await page.waitForTimeout(200);
    await emailBadge.click({ modifiers: ['Shift'] });
    await page.waitForTimeout(200);

    const state = await getWinState(page, 'sample1');
    expect(state.sortCols).toEqual([
      { col: 'name', dir: 'asc' },
      { col: 'email', dir: 'asc' },
    ]);
  });

  test('sort badge click does not change selectedCells', async ({ page }) => {
    // First select a column
    const header = page.locator('.subwindow table thead th').nth(2); // email
    await header.click();
    await page.waitForTimeout(100);

    const before = await getWinState(page, 'sample1');
    expect(before.selectedCells.length).toBe(10);

    // Click sort badge on a different column
    const nameBadge = page.locator('.subwindow table thead th').nth(1).locator('.col-badge-sort');
    await nameBadge.click();
    await page.waitForTimeout(200);

    const after = await getWinState(page, 'sample1');
    expect(after.selectedCells).toEqual(before.selectedCells);
  });

  // --- selectedCol Lifecycle ---

  test('cell focusin clears selectedCol', async ({ page }) => {
    const header = page.locator('.subwindow table thead th').nth(1);
    await header.click();
    await page.waitForTimeout(100);

    let state = await getWinState(page, 'sample1');
    expect(state.selectedCol).toBe('name');

    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.waitForTimeout(100);

    state = await getWinState(page, 'sample1');
    expect(state.selectedCol).toBeNull();
  });

  test('row selection clears selectedCol', async ({ page }) => {
    const header = page.locator('.subwindow table thead th').nth(1);
    await header.click();
    await page.waitForTimeout(100);

    const rowNum = page.locator('.subwindow table tbody td.row-num').first();
    await rowNum.click();
    await page.waitForTimeout(100);

    const state = await getWinState(page, 'sample1');
    expect(state.selectedCol).toBeNull();
  });

  test('selectAllCells clears selectedCol', async ({ page }) => {
    const header = page.locator('.subwindow table thead th').nth(1);
    await header.click();
    await page.waitForTimeout(100);

    // Focus a cell so Ctrl+A targets the table
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.keyboard.press('Control+a');
    await page.waitForTimeout(100);

    const state = await getWinState(page, 'sample1');
    expect(state.selectedCol).toBeNull();
  });

  test('selectNoneCells clears selectedCol', async ({ page }) => {
    const header = page.locator('.subwindow table thead th').nth(1);
    await header.click();
    await page.waitForTimeout(100);

    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.keyboard.press('Control+Shift+a');
    await page.waitForTimeout(100);

    const state = await getWinState(page, 'sample1');
    expect(state.selectedCol).toBeNull();
  });

  // --- Ctrl+Arrow Gating ---

  test('Ctrl+Arrow does nothing after cell click', async ({ page }) => {
    const cols0 = await page.evaluate(() => [...app._test.tables['sample1'].columns]);

    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.keyboard.press('Control+ArrowRight');
    await page.waitForTimeout(200);

    const cols1 = await page.evaluate(() => [...app._test.tables['sample1'].columns]);
    expect(cols1).toEqual(cols0);
  });

  test('Ctrl+Arrow does nothing after row selection', async ({ page }) => {
    const cols0 = await page.evaluate(() => [...app._test.tables['sample1'].columns]);

    const rowNum = page.locator('.subwindow table tbody td.row-num').first();
    await rowNum.click();
    await page.waitForTimeout(100);
    await page.evaluate(() => document.body.focus());
    await page.keyboard.press('Control+ArrowRight');
    await page.waitForTimeout(200);

    const cols1 = await page.evaluate(() => [...app._test.tables['sample1'].columns]);
    expect(cols1).toEqual(cols0);
  });

  test('Ctrl+Arrow works after header click', async ({ page }) => {
    const header = page.locator('.subwindow table thead th').nth(1);
    await header.click();
    await page.evaluate(() => document.body.focus());
    await page.keyboard.press('Control+ArrowRight');
    await page.waitForTimeout(200);

    const cols = await page.evaluate(() => [...app._test.tables['sample1'].columns]);
    expect(cols).toEqual(['email', 'name', 'member_since']);
  });

  // --- Copy with Column Selection ---

  test('copy after header click includes header row', async ({ context, page }) => {
    await context.grantPermissions(['clipboard-read', 'clipboard-write']);

    const header = page.locator('.subwindow table thead th').nth(1);
    await header.click();
    await page.waitForTimeout(100);

    const state = await getWinState(page, 'sample1');
    expect(state.copyWithHeader).toBe(true);
  });

  test('column selection clipboard copy produces correct TSV', async ({ context, page }) => {
    await context.grantPermissions(['clipboard-read', 'clipboard-write']);

    const header = page.locator('.subwindow table thead th').nth(1);
    await header.click();
    await page.waitForTimeout(100);

    await page.evaluate(async () => {
      const win = app._test.windows[0];
      await app._test.copySelectedCells(win);
    });
    await page.waitForTimeout(200);

    const clipText = await page.evaluate(() => navigator.clipboard.readText());
    const lines = clipText.trim().split('\n');
    expect(lines[0]).toBe('name');
    expect(lines.length).toBe(11); // header + 10 data rows
  });

  // --- selectColumns function ---

  test('selectColumns with single column selects all cells in that column', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      app._test.selectColumns(win, 1, 1); // email
    });
    await page.waitForTimeout(100);

    const state = await getWinState(page, 'sample1');
    expect(state.selectedCells.length).toBe(10);
    expect(state.selectedCells.every(c => c.endsWith(':email'))).toBe(true);
    expect(state.copyWithHeader).toBe(true);
  });

  test('selectColumns with range selects all cells across multiple columns', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      app._test.selectColumns(win, 0, 2); // name, email, member_since
    });
    await page.waitForTimeout(100);

    const state = await getWinState(page, 'sample1');
    expect(state.selectedCells.length).toBe(30);
    const cols = new Set(state.selectedCells.map(c => c.split(':')[1]));
    expect([...cols].sort()).toEqual(['email', 'member_since', 'name']);
    expect(state.anchorCell).toEqual({ rownum: 1, col: 'name' });
  });
});
