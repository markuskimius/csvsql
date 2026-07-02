const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, executeSQL } = require('../helpers');

async function getSelection(page, tableName) {
  return page.evaluate((name) => {
    const w = app._test.windows.find(w => w.tableName === name);
    if (!w) return null;
    return {
      cells: [...(w.selectedCells || [])].sort(),
      anchor: w.anchorCell,
    };
  }, tableName);
}

test.describe('Cell row/column highlight', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('clicking a cell highlights its row and column', async ({ page }) => {
    const firstCell = page.locator('.subwindow table tbody td.data-cell').first();
    await firstCell.click();

    // Row highlight on the TR, col highlight on the TH and every TD in that column
    const tr = firstCell.locator('xpath=..');
    await expect(tr).toHaveClass(/row-highlight/);

    const th = page.locator('.subwindow table thead th').nth(1);
    await expect(th).toHaveClass(/col-highlight/);

    await expect(firstCell).toHaveClass(/cell-selected/);

    const sel = await getSelection(page, 'sample1');
    expect(sel.cells.length).toBe(1);
    expect(sel.anchor).toEqual({ rownum: 1, col: 'name' });
  });

  test('Shift+Right extends selection by one cell', async ({ page }) => {
    const firstCell = page.locator('.subwindow table tbody td.data-cell').first();
    await firstCell.click();

    await page.keyboard.press('Shift+ArrowRight');
    await page.waitForTimeout(150);

    const sel = await getSelection(page, 'sample1');
    expect(sel.cells).toEqual(['1:email', '1:name']);
    expect(sel.anchor).toEqual({ rownum: 1, col: 'name' });

    // Both columns highlighted
    const ths = page.locator('.subwindow table thead th.col-highlight');
    await expect(ths).toHaveCount(2);
  });

  test('Shift+Down extends selection across rows', async ({ page }) => {
    const firstCell = page.locator('.subwindow table tbody td.data-cell').first();
    await firstCell.click();

    await page.keyboard.press('Shift+ArrowDown');
    await page.keyboard.press('Shift+ArrowDown');
    await page.waitForTimeout(150);

    const sel = await getSelection(page, 'sample1');
    expect(sel.cells).toEqual(['1:name', '2:name', '3:name']);

    const highlightedRows = page.locator('.subwindow table tbody tr.row-highlight');
    await expect(highlightedRows).toHaveCount(3);
  });

  test('shrinking the rectangle removes cells from selection', async ({ page }) => {
    const firstCell = page.locator('.subwindow table tbody td.data-cell').first();
    await firstCell.click();

    await page.keyboard.press('Shift+ArrowRight');
    await page.keyboard.press('Shift+ArrowDown');
    await page.waitForTimeout(100);
    let sel = await getSelection(page, 'sample1');
    expect(sel.cells.length).toBe(4); // 2x2 rectangle

    // Shrink by pulling lead back up
    await page.keyboard.press('Shift+ArrowUp');
    await page.waitForTimeout(100);
    sel = await getSelection(page, 'sample1');
    expect(sel.cells.length).toBe(2); // 1x2
    expect(sel.cells).toEqual(['1:email', '1:name']);
  });

  test('Esc clears the selection', async ({ page }) => {
    const firstCell = page.locator('.subwindow table tbody td.data-cell').first();
    await firstCell.click();
    await page.keyboard.press('Shift+ArrowDown');
    await page.waitForTimeout(100);

    await page.keyboard.press('Escape');
    await page.waitForTimeout(100);

    const sel = await getSelection(page, 'sample1');
    expect(sel.cells).toEqual([]);
    expect(sel.anchor).toBeNull();

    const highlighted = page.locator('.subwindow table tbody tr.row-highlight');
    await expect(highlighted).toHaveCount(0);
  });

  test('selection survives column reorder (keyed by name, not index)', async ({ page }) => {
    // Click header to select entire "name" column, then Ctrl+Right to swap with email
    const nameHeader = page.locator('.subwindow table thead th').nth(1);
    await nameHeader.click();
    await page.waitForTimeout(100);

    const selBefore = await getSelection(page, 'sample1');
    // All 10 rows of name column selected
    expect(selBefore.cells.length).toBe(10);
    expect(selBefore.cells.every(c => c.endsWith(':name'))).toBe(true);

    await page.evaluate(() => document.body.focus());
    await page.keyboard.press('Control+ArrowRight');
    await page.waitForTimeout(200);

    const sel = await getSelection(page, 'sample1');
    // Selection keys reference column names — all name cells survive reorder
    expect(sel.cells.length).toBe(10);
    expect(sel.cells.every(c => c.endsWith(':name'))).toBe(true);

    // name TH should still have col-highlight (by name, via applyCellHighlights)
    const highlightedTh = page.locator('.subwindow table thead th.col-highlight');
    await expect(highlightedTh).toHaveCount(1);
  });

  test('Shift+click extends selection from existing anchor', async ({ page }) => {
    const firstCell = page.locator('.subwindow table tbody td.data-cell').first();
    await firstCell.click();

    // Shift-click a cell 2 rows down, 1 col over (email of row 3)
    const target = page.locator('.subwindow table tbody tr[data-display-idx="2"] td.data-cell').nth(1);
    await target.click({ modifiers: ['Shift'] });
    await page.waitForTimeout(100);

    const sel = await getSelection(page, 'sample1');
    expect(sel.cells).toEqual(['1:email', '1:name', '2:email', '2:name', '3:email', '3:name']);
    // Anchor stays on the originally clicked cell
    expect(sel.anchor).toEqual({ rownum: 1, col: 'name' });
  });

  test('mouse drag sweeps a rectangle of cells', async ({ page }) => {
    const start = page.locator('.subwindow table tbody tr[data-display-idx="0"] td.data-cell').nth(0); // row 1 name
    const end = page.locator('.subwindow table tbody tr[data-display-idx="2"] td.data-cell').nth(1); // row 3 email

    const sBox = await start.boundingBox();
    const eBox = await end.boundingBox();

    await page.mouse.move(sBox.x + sBox.width / 2, sBox.y + sBox.height / 2);
    await page.mouse.down();
    await page.mouse.move(eBox.x + eBox.width / 2, eBox.y + eBox.height / 2, { steps: 10 });
    await page.mouse.up();
    await page.waitForTimeout(100);

    const sel = await getSelection(page, 'sample1');
    expect(sel.cells).toEqual(['1:email', '1:name', '2:email', '2:name', '3:email', '3:name']);
    expect(sel.anchor).toEqual({ rownum: 1, col: 'name' });

    // After drag, Shift+Down should extend from the drag-end cell (not anchor)
    await page.keyboard.press('Shift+ArrowDown');
    await page.waitForTimeout(100);
    const sel2 = await getSelection(page, 'sample1');
    expect(sel2.cells.length).toBe(8); // now 4 rows × 2 cols
  });

  test('Ctrl+Arrow reorder requires column header selection', async ({ page }) => {
    const cols0 = await page.evaluate(() => [...app._test.tables['sample1'].columns]);

    // Select a cell — Ctrl+Right should NOT reorder
    const firstCell = page.locator('.subwindow table tbody td.data-cell').first();
    await firstCell.click();
    await page.keyboard.press('Control+ArrowRight');
    await page.waitForTimeout(150);
    const cols1 = await page.evaluate(() => [...app._test.tables['sample1'].columns]);
    expect(cols1).toEqual(cols0);

    // Select a row — Ctrl+Right should NOT reorder
    await page.locator('.subwindow table tbody td.row-num').first().click();
    await page.waitForTimeout(100);
    await page.evaluate(() => document.body.focus());
    await page.keyboard.press('Control+ArrowRight');
    await page.waitForTimeout(150);
    const cols2 = await page.evaluate(() => [...app._test.tables['sample1'].columns]);
    expect(cols2).toEqual(cols0);

    // Select column via header — Ctrl+Right SHOULD reorder
    const firstHeader = page.locator('.subwindow table thead th').nth(1);
    await firstHeader.click();
    await page.evaluate(() => document.body.focus());
    await page.keyboard.press('Control+ArrowRight');
    await page.waitForTimeout(200);
    const cols3 = await page.evaluate(() => [...app._test.tables['sample1'].columns]);
    expect(cols3).toEqual(['email', 'name', 'member_since']);
  });

  test('Ctrl+Arrow at edge is a no-op', async ({ page }) => {
    const firstCell = page.locator('.subwindow table tbody td.data-cell').first();
    await firstCell.click();
    // Selection is just column "name" (leftmost). Ctrl+Left should do nothing.
    await page.keyboard.press('Control+ArrowLeft');
    await page.waitForTimeout(150);
    const cols = await page.evaluate(() => [...app._test.tables['sample1'].columns]);
    expect(cols).toEqual(['name', 'email', 'member_since']);
  });

  test('arrow key with no selection focuses a cell in the middle of view', async ({ page }) => {
    // Activate the window (click its title bar so a cell is NOT focused).
    await page.locator('.subwindow .win-titlebar').first().click();
    await page.waitForTimeout(50);

    // Nothing is selected yet.
    let sel = await getSelection(page, 'sample1');
    expect(sel.anchor).toBeNull();

    await page.keyboard.press('ArrowDown');
    await page.waitForTimeout(100);

    sel = await getSelection(page, 'sample1');
    expect(sel.anchor).not.toBeNull();
    // sample1 has 10 rows, 2 cols. Middle of view is around row 5 ± a few.
    expect(sel.anchor.rownum).toBeGreaterThanOrEqual(3);
    expect(sel.anchor.rownum).toBeLessThanOrEqual(7);
  });

  test('plain arrow moves single-cell selection (collapses rect)', async ({ page }) => {
    const firstCell = page.locator('.subwindow table tbody td.data-cell').first();
    await firstCell.click();
    // Extend to a 2x2 rectangle first.
    await page.keyboard.press('Shift+ArrowRight');
    await page.keyboard.press('Shift+ArrowDown');
    await page.waitForTimeout(100);
    let sel = await getSelection(page, 'sample1');
    expect(sel.cells.length).toBe(4);

    // Plain ArrowDown: collapse to a single cell one row below the lead (rownum 2, email).
    await page.keyboard.press('ArrowDown');
    await page.waitForTimeout(100);

    sel = await getSelection(page, 'sample1');
    expect(sel.cells).toEqual(['3:email']);
    expect(sel.anchor).toEqual({ rownum: 3, col: 'email' });
  });

  test('left/right arrows in edit mode move the caret, not the selection', async ({ page }) => {
    const firstCell = page.locator('.subwindow table tbody td.data-cell').first();
    await firstCell.click();
    await page.keyboard.press('F2');
    await page.waitForTimeout(50);

    // In edit mode, left/right arrows move the caret; anchor unchanged.
    await page.keyboard.press('ArrowLeft');
    await page.keyboard.press('ArrowRight');
    await page.waitForTimeout(50);

    const sel = await getSelection(page, 'sample1');
    expect(sel.anchor).toEqual({ rownum: 1, col: 'name' });
  });

  test('down arrow in edit mode commits the edit and moves selection down', async ({ page }) => {
    const firstCell = page.locator('.subwindow table tbody td.data-cell').first();
    await firstCell.click();
    await page.keyboard.press('F2');
    await page.waitForTimeout(50);

    await page.keyboard.press('ArrowDown');
    await page.waitForTimeout(100);

    const sel = await getSelection(page, 'sample1');
    expect(sel.anchor).toEqual({ rownum: 2, col: 'name' });
  });

  test('clicking a different cell resets anchor and selection', async ({ page }) => {
    const firstCell = page.locator('.subwindow table tbody td.data-cell').first();
    await firstCell.click();
    await page.keyboard.press('Control+Shift+ArrowDown');
    await page.keyboard.press('Control+Shift+ArrowDown');
    await page.waitForTimeout(100);

    // Click a totally different cell
    const target = page.locator('.subwindow table tbody tr[data-display-idx="4"] td.data-cell').nth(1);
    await target.click();
    await page.waitForTimeout(100);

    const sel = await getSelection(page, 'sample1');
    expect(sel.cells).toEqual(['5:email']);
    expect(sel.anchor).toEqual({ rownum: 5, col: 'email' });
  });

  test('Select All preserves horizontal scroll position', async ({ page }) => {
    // Create a wide table so the container has a horizontal scrollbar
    await executeSQL(page, "SELECT 'a' AS c1, 'b' AS c2, 'c' AS c3, 'd' AS c4, 'e' AS c5, 'f' AS c6, 'g' AS c7, 'h' AS c8, 'i' AS c9, 'j' AS c10, 'k' AS c11, 'l' AS c12, 'm' AS c13, 'n' AS c14, 'o' AS c15");
    const qName = await page.evaluate(() => {
      const w = app._test.windows.find(w => w.isQuery);
      return w ? w.tableName : null;
    });
    expect(qName).toBeTruthy();

    // Make the window narrow enough to need horizontal scrolling
    await page.evaluate((name) => {
      const win = app._test.windows.find(w => w.tableName === name);
      const el = document.getElementById('win-' + win.id);
      el.style.width = '400px';
      win.width = 400;
    }, qName);

    // Get the table container and scroll it right
    const container = page.locator('.subwindow.active .table-container');
    await container.evaluate(el => { el.scrollLeft = 200; });
    await page.waitForTimeout(100);

    const scrollBefore = await container.evaluate(el => el.scrollLeft);
    expect(scrollBefore).toBeGreaterThan(0);

    // Click a visible cell to give focus to the table
    const cell = page.locator('.subwindow.active table tbody td.data-cell').first();
    await cell.click();
    await page.waitForTimeout(50);

    // Re-scroll since clicking may have changed scroll
    await container.evaluate(el => { el.scrollLeft = 200; });
    await page.waitForTimeout(100);
    const scrollBeforeSelectAll = await container.evaluate(el => el.scrollLeft);
    expect(scrollBeforeSelectAll).toBeGreaterThan(0);

    // Select All via Ctrl+A
    await page.keyboard.press('Control+a');
    await page.waitForTimeout(150);

    // Verify all cells are selected
    const sel = await page.evaluate((name) => {
      const w = app._test.windows.find(w => w.tableName === name);
      return w ? w.selectedCells.size : 0;
    }, qName);
    expect(sel).toBe(15); // 1 row × 15 columns

    // Verify horizontal scroll position was preserved
    const scrollAfter = await container.evaluate(el => el.scrollLeft);
    expect(scrollAfter).toBe(scrollBeforeSelectAll);
  });

  test('Select All via _test API preserves horizontal scroll position', async ({ page }) => {
    await executeSQL(page, "SELECT 'a' AS c1, 'b' AS c2, 'c' AS c3, 'd' AS c4, 'e' AS c5, 'f' AS c6, 'g' AS c7, 'h' AS c8, 'i' AS c9, 'j' AS c10, 'k' AS c11, 'l' AS c12, 'm' AS c13, 'n' AS c14, 'o' AS c15");
    const qName = await page.evaluate(() => {
      const w = app._test.windows.find(w => w.isQuery);
      return w ? w.tableName : null;
    });
    expect(qName).toBeTruthy();

    await page.evaluate((name) => {
      const win = app._test.windows.find(w => w.tableName === name);
      const el = document.getElementById('win-' + win.id);
      el.style.width = '400px';
      win.width = 400;
    }, qName);

    const container = page.locator('.subwindow.active .table-container');
    await container.evaluate(el => { el.scrollLeft = 200; });
    await page.waitForTimeout(100);
    const scrollBefore = await container.evaluate(el => el.scrollLeft);
    expect(scrollBefore).toBeGreaterThan(0);

    // Call selectAllCells directly to test the code path without click side effects
    await page.evaluate((name) => {
      const win = app._test.windows.find(w => w.tableName === name);
      app._test.selectAllCells(win);
    }, qName);
    await page.waitForTimeout(150);

    const sel = await page.evaluate((name) => {
      const w = app._test.windows.find(w => w.tableName === name);
      return w ? w.selectedCells.size : 0;
    }, qName);
    expect(sel).toBe(15);

    const scrollAfter = await container.evaluate(el => el.scrollLeft);
    expect(scrollAfter).toBe(scrollBefore);
  });

  test('Select All via Edit menu preserves horizontal scroll position', async ({ page }) => {
    await executeSQL(page, "SELECT 'a' AS c1, 'b' AS c2, 'c' AS c3, 'd' AS c4, 'e' AS c5, 'f' AS c6, 'g' AS c7, 'h' AS c8, 'i' AS c9, 'j' AS c10, 'k' AS c11, 'l' AS c12, 'm' AS c13, 'n' AS c14, 'o' AS c15");
    const qName = await page.evaluate(() => {
      const w = app._test.windows.find(w => w.isQuery);
      return w ? w.tableName : null;
    });
    expect(qName).toBeTruthy();

    await page.evaluate((name) => {
      const win = app._test.windows.find(w => w.tableName === name);
      const el = document.getElementById('win-' + win.id);
      el.style.width = '400px';
      win.width = 400;
    }, qName);

    const container = page.locator('.subwindow.active .table-container');
    await container.evaluate(el => { el.scrollLeft = 200; });
    await page.waitForTimeout(100);
    const scrollBefore = await container.evaluate(el => el.scrollLeft);
    expect(scrollBefore).toBeGreaterThan(0);

    // Use Edit menu Select All button
    await page.locator('#menu-edit .menu-label').click();
    await page.waitForTimeout(100);
    await page.locator('#btn-select-all').click();
    await page.waitForTimeout(150);

    const sel = await page.evaluate((name) => {
      const w = app._test.windows.find(w => w.tableName === name);
      return w ? w.selectedCells.size : 0;
    }, qName);
    expect(sel).toBe(15);

    const scrollAfter = await container.evaluate(el => el.scrollLeft);
    expect(scrollAfter).toBe(scrollBefore);
  });

  test('Select None preserves horizontal scroll position', async ({ page }) => {
    await executeSQL(page, "SELECT 'a' AS c1, 'b' AS c2, 'c' AS c3, 'd' AS c4, 'e' AS c5, 'f' AS c6, 'g' AS c7, 'h' AS c8, 'i' AS c9, 'j' AS c10, 'k' AS c11, 'l' AS c12, 'm' AS c13, 'n' AS c14, 'o' AS c15");
    const qName = await page.evaluate(() => {
      const w = app._test.windows.find(w => w.isQuery);
      return w ? w.tableName : null;
    });
    expect(qName).toBeTruthy();

    await page.evaluate((name) => {
      const win = app._test.windows.find(w => w.tableName === name);
      const el = document.getElementById('win-' + win.id);
      el.style.width = '400px';
      win.width = 400;
    }, qName);

    // Select all first
    await page.evaluate((name) => {
      const win = app._test.windows.find(w => w.tableName === name);
      app._test.selectAllCells(win);
    }, qName);
    await page.waitForTimeout(100);

    const container = page.locator('.subwindow.active .table-container');
    await container.evaluate(el => { el.scrollLeft = 200; });
    await page.waitForTimeout(100);
    const scrollBefore = await container.evaluate(el => el.scrollLeft);
    expect(scrollBefore).toBeGreaterThan(0);

    // Deselect via selectNoneCells (no cell focused in this context)
    await page.evaluate((name) => {
      const win = app._test.windows.find(w => w.tableName === name);
      app._test.selectNoneCells(win);
    }, qName);
    await page.waitForTimeout(150);

    const sel = await page.evaluate((name) => {
      const w = app._test.windows.find(w => w.tableName === name);
      return w ? w.selectedCells.size : 0;
    }, qName);
    expect(sel).toBe(0);

    const scrollAfter = await container.evaluate(el => el.scrollLeft);
    expect(scrollAfter).toBe(scrollBefore);
  });
});
