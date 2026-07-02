const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, getTableData } = require('../helpers');

// ---------------------------------------------------------------------------
// Edit Mode Navigation
// ---------------------------------------------------------------------------

test.describe('Edit mode navigation', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('Tab in edit mode moves to the next cell and keeps edit mode active', async ({ page }) => {
    const cell = page.locator('.subwindow table tbody tr:first-child td.data-cell').first();
    await cell.click();
    await page.keyboard.press('F2');
    await expect(cell).toHaveAttribute('contenteditable', 'true');

    // Tab should move to the next cell in the same row
    await page.keyboard.press('Tab');
    await page.waitForTimeout(200);

    const nextCell = page.locator('.subwindow table tbody tr:first-child td.data-cell').nth(1);
    await expect(nextCell).toHaveAttribute('contenteditable', 'true');
    // The original cell should no longer be editable
    await expect(cell).not.toHaveAttribute('contenteditable', 'true');
  });

  test('Shift+Tab in edit mode moves to the previous cell and keeps edit mode active', async ({ page }) => {
    // Start on the second cell
    const secondCell = page.locator('.subwindow table tbody tr:first-child td.data-cell').nth(1);
    await secondCell.click();
    await page.keyboard.press('F2');
    await expect(secondCell).toHaveAttribute('contenteditable', 'true');

    // Shift+Tab should move to the previous cell
    await page.keyboard.press('Shift+Tab');
    await page.waitForTimeout(200);

    const firstCell = page.locator('.subwindow table tbody tr:first-child td.data-cell').first();
    await expect(firstCell).toHaveAttribute('contenteditable', 'true');
    await expect(secondCell).not.toHaveAttribute('contenteditable', 'true');
  });

  test('Enter in edit mode saves and moves to next row at the start column', async ({ page }) => {
    // Enter edit mode on the first cell (column 0)
    const cell = page.locator('.subwindow table tbody tr:first-child td.data-cell').first();
    await cell.click();
    await page.keyboard.press('F2');
    await expect(cell).toHaveAttribute('contenteditable', 'true');

    // Press Enter to save and move down
    await page.keyboard.press('Enter');
    await page.waitForTimeout(200);

    // The cell in the next row at column 0 should be in edit mode
    const nextRowCell = page.locator('.subwindow table tbody tr:nth-child(2) td.data-cell').first();
    await expect(nextRowCell).toHaveAttribute('contenteditable', 'true');
    // Original cell should not be editable
    await expect(cell).not.toHaveAttribute('contenteditable', 'true');
  });

  test('Enter navigates to _editStartColIdx, not the current column', async ({ page }) => {
    // Enter edit mode on column 0
    const firstCell = page.locator('.subwindow table tbody tr:first-child td.data-cell').first();
    await firstCell.click();
    await page.keyboard.press('F2');

    // Tab to column 1
    await page.keyboard.press('Tab');
    await page.waitForTimeout(100);

    // Tab to column 2
    await page.keyboard.press('Tab');
    await page.waitForTimeout(100);

    // Now in column 2 (member_since). Press Enter — should go to column 0 on next row
    await page.keyboard.press('Enter');
    await page.waitForTimeout(200);

    // The target should be row 2, column 0 (the start column)
    const targetCell = page.locator('.subwindow table tbody tr:nth-child(2) td.data-cell').first();
    await expect(targetCell).toHaveAttribute('contenteditable', 'true');

    // Verify column 2 of row 2 is NOT in edit mode
    const col2Row2 = page.locator('.subwindow table tbody tr:nth-child(2) td.data-cell').nth(2);
    await expect(col2Row2).not.toHaveAttribute('contenteditable', 'true');
  });

  test('_editStartColIdx is set when entering edit mode via F2', async ({ page }) => {
    const secondCell = page.locator('.subwindow table tbody tr:first-child td.data-cell').nth(1);
    await secondCell.click();
    await page.keyboard.press('F2');

    const startColIdx = await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      return win._editStartColIdx;
    });
    expect(startColIdx).toBe(1);
  });

  test('Enter in select mode moves down (does not enter edit mode)', async ({ page }) => {
    const cell = page.locator('.subwindow table tbody tr:first-child td.data-cell').first();
    await cell.click();
    await expect(cell).not.toHaveAttribute('contenteditable', 'true');

    await page.keyboard.press('Enter');
    await page.waitForTimeout(200);

    // Should NOT enter edit mode — the cell should still not be contenteditable
    await expect(cell).not.toHaveAttribute('contenteditable', 'true');

    // Focus should have moved to the second row
    const focused = await page.evaluate(() => {
      const el = document.activeElement;
      return el && el.classList.contains('data-cell')
        ? parseInt(el.parentElement.dataset.displayIdx, 10)
        : null;
    });
    expect(focused).toBe(1);
  });

  test('_editStartColIdx is set when entering edit mode via i key', async ({ page }) => {
    const firstCell = page.locator('.subwindow table tbody tr:first-child td.data-cell').first();
    await firstCell.click();
    await page.keyboard.press('i');

    const startColIdx = await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      return win._editStartColIdx;
    });
    expect(startColIdx).toBe(0);
  });

  test('_editStartColIdx is set when entering edit mode via Ctrl+U', async ({ page }) => {
    const secondCell = page.locator('.subwindow table tbody tr:first-child td.data-cell').nth(1);
    await secondCell.click();
    await page.keyboard.press('Control+u');

    const startColIdx = await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      return win._editStartColIdx;
    });
    expect(startColIdx).toBe(1);
  });

  test('_editStartColIdx is set when entering edit mode via Ctrl+click', async ({ page }) => {
    const thirdCell = page.locator('.subwindow table tbody tr:first-child td.data-cell').nth(2);
    await thirdCell.click({ modifiers: ['Control'] });

    const startColIdx = await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      return win._editStartColIdx;
    });
    expect(startColIdx).toBe(2);
  });

  test('Tab preserves _editStartColIdx across multiple Tab presses', async ({ page }) => {
    // Enter edit on column 0
    const firstCell = page.locator('.subwindow table tbody tr:first-child td.data-cell').first();
    await firstCell.click();
    await page.keyboard.press('F2');

    // Tab to column 1
    await page.keyboard.press('Tab');
    await page.waitForTimeout(100);

    // _editStartColIdx should still be 0
    const startColIdx = await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      return win._editStartColIdx;
    });
    expect(startColIdx).toBe(0);
  });

  test('Tab at last cell in row does not crash', async ({ page }) => {
    // Focus the last cell in the first row
    const lastCell = page.locator('.subwindow table tbody tr:first-child td.data-cell').last();
    await lastCell.click();
    await page.keyboard.press('F2');
    await expect(lastCell).toHaveAttribute('contenteditable', 'true');

    // Tab at the end should not crash; the last cell stays in edit mode or exits gracefully
    await page.keyboard.press('Tab');
    await page.waitForTimeout(200);

    // Verify no crash occurred by checking the page is still functional
    const data = await getTableData(page, 'sample1');
    expect(data.rows.length).toBe(10);
  });

  test('Shift+Tab at first cell in row does not crash', async ({ page }) => {
    const firstCell = page.locator('.subwindow table tbody tr:first-child td.data-cell').first();
    await firstCell.click();
    await page.keyboard.press('F2');
    await expect(firstCell).toHaveAttribute('contenteditable', 'true');

    // Shift+Tab at the beginning should not crash
    await page.keyboard.press('Shift+Tab');
    await page.waitForTimeout(200);

    const data = await getTableData(page, 'sample1');
    expect(data.rows.length).toBe(10);
  });

  test('Enter on last row exits edit mode and keeps cell focused', async ({ page }) => {
    const lastRowCell = page.locator('.subwindow table tbody tr:last-child td.data-cell').first();
    await lastRowCell.click();
    await page.keyboard.press('F2');
    await expect(lastRowCell).toHaveAttribute('contenteditable', 'true');

    await page.keyboard.press('Enter');
    await page.waitForTimeout(200);

    // Should exit edit mode (no contenteditable)
    const isEditable = await page.evaluate(() => {
      const el = document.activeElement;
      return el && el.getAttribute('contenteditable') === 'true';
    });
    expect(isEditable).toBe(false);

    // Should still have a data cell focused (navigable)
    const hasFocus = await page.evaluate(() => {
      const el = document.activeElement;
      return el && el.classList.contains('data-cell');
    });
    expect(hasFocus).toBe(true);

    // Arrow keys should still work from here
    await page.keyboard.press('ArrowUp');
    await page.waitForTimeout(200);
    const movedUp = await page.evaluate(() => {
      const el = document.activeElement;
      return el && el.classList.contains('data-cell')
        ? parseInt(el.parentElement.dataset.displayIdx, 10)
        : null;
    });
    expect(movedUp).toBeLessThan(9);
  });

  test('ArrowDown in edit mode commits and moves to next row in nav mode', async ({ page }) => {
    const cell = page.locator('.subwindow table tbody tr:first-child td.data-cell').first();
    await cell.click();
    await page.keyboard.press('F2');
    await expect(cell).toHaveAttribute('contenteditable', 'true');
    await cell.fill('ARROW_DOWN_TEST');

    await page.keyboard.press('ArrowDown');
    await page.waitForTimeout(200);

    // Should have committed the edit
    const data = await getTableData(page, 'sample1');
    expect(data.rows[0].name).toBe('ARROW_DOWN_TEST');

    // Should be on the next row
    const focused = await page.evaluate(() => {
      const el = document.activeElement;
      return el && el.classList.contains('data-cell')
        ? parseInt(el.parentElement.dataset.displayIdx, 10)
        : null;
    });
    expect(focused).toBe(1);

    // Should NOT be in edit mode
    const isEditable = await page.evaluate(() => {
      const el = document.activeElement;
      return el && el.getAttribute('contenteditable') === 'true';
    });
    expect(isEditable).toBe(false);
  });

  test('ArrowUp in edit mode commits and moves to previous row in nav mode', async ({ page }) => {
    const secondRowCell = page.locator('.subwindow table tbody tr:nth-child(2) td.data-cell').first();
    await secondRowCell.click();
    await page.keyboard.press('F2');
    await expect(secondRowCell).toHaveAttribute('contenteditable', 'true');
    await secondRowCell.fill('ARROW_UP_TEST');

    await page.keyboard.press('ArrowUp');
    await page.waitForTimeout(200);

    // Should have committed the edit
    const data = await getTableData(page, 'sample1');
    expect(data.rows[1].name).toBe('ARROW_UP_TEST');

    // Should be on the first row
    const focused = await page.evaluate(() => {
      const el = document.activeElement;
      return el && el.classList.contains('data-cell')
        ? parseInt(el.parentElement.dataset.displayIdx, 10)
        : null;
    });
    expect(focused).toBe(0);

    // Should NOT be in edit mode
    const isEditable = await page.evaluate(() => {
      const el = document.activeElement;
      return el && el.getAttribute('contenteditable') === 'true';
    });
    expect(isEditable).toBe(false);
  });

  test('ArrowDown on last row in edit mode commits and stays on last row', async ({ page }) => {
    const lastRowCell = page.locator('.subwindow table tbody tr:last-child td.data-cell').first();
    await lastRowCell.click();
    await page.keyboard.press('F2');

    await page.keyboard.press('ArrowDown');
    await page.waitForTimeout(200);

    // Should still be on the last row (display index 9)
    const focused = await page.evaluate(() => {
      const el = document.activeElement;
      return el && el.classList.contains('data-cell')
        ? parseInt(el.parentElement.dataset.displayIdx, 10)
        : null;
    });
    expect(focused).toBe(9);
  });

  test('ArrowUp on first row in edit mode commits and stays on first row', async ({ page }) => {
    const firstCell = page.locator('.subwindow table tbody tr:first-child td.data-cell').first();
    await firstCell.click();
    await page.keyboard.press('F2');

    await page.keyboard.press('ArrowUp');
    await page.waitForTimeout(200);

    const focused = await page.evaluate(() => {
      const el = document.activeElement;
      return el && el.classList.contains('data-cell')
        ? parseInt(el.parentElement.dataset.displayIdx, 10)
        : null;
    });
    expect(focused).toBe(0);
  });

  test('ArrowDown in edit mode stays in same column', async ({ page }) => {
    // Start editing in column 1
    const cell = page.locator('.subwindow table tbody tr:first-child td.data-cell').nth(1);
    await cell.click();
    await page.keyboard.press('F2');

    await page.keyboard.press('ArrowDown');
    await page.waitForTimeout(200);

    const focused = await page.evaluate(() => {
      const el = document.activeElement;
      return el && el.classList.contains('data-cell')
        ? { row: parseInt(el.parentElement.dataset.displayIdx, 10), col: parseInt(el.dataset.colIdx, 10) }
        : null;
    });
    expect(focused).toEqual({ row: 1, col: 1 });
  });

  test('Enter saves the edited value before moving down', async ({ page }) => {
    const cell = page.locator('.subwindow table tbody tr:first-child td.data-cell').first();
    await cell.click();
    await page.keyboard.press('F2');
    await cell.fill('EDITED_VIA_ENTER');
    await page.keyboard.press('Enter');
    await page.waitForTimeout(300);

    const data = await getTableData(page, 'sample1');
    expect(data.rows[0].name).toBe('EDITED_VIA_ENTER');
  });

  test('Tab in select mode moves to next cell (not cycles windows)', async ({ page }) => {
    const cell = page.locator('.subwindow table tbody tr:first-child td.data-cell').first();
    await cell.click();
    await expect(cell).not.toHaveAttribute('contenteditable', 'true');

    await page.keyboard.press('Tab');
    await page.waitForTimeout(200);

    // Focus should have moved to the next cell in the same row
    const focused = await page.evaluate(() => {
      const el = document.activeElement;
      return el && el.classList.contains('data-cell')
        ? parseInt(el.dataset.colIdx, 10)
        : null;
    });
    expect(focused).toBe(1);
    // Should still be in select mode
    const nextCell = page.locator('.subwindow table tbody tr:first-child td.data-cell').nth(1);
    await expect(nextCell).not.toHaveAttribute('contenteditable', 'true');
  });

  test('Shift+Tab in select mode moves to previous cell', async ({ page }) => {
    const secondCell = page.locator('.subwindow table tbody tr:first-child td.data-cell').nth(1);
    await secondCell.click();
    await expect(secondCell).not.toHaveAttribute('contenteditable', 'true');

    await page.keyboard.press('Shift+Tab');
    await page.waitForTimeout(200);

    const focused = await page.evaluate(() => {
      const el = document.activeElement;
      return el && el.classList.contains('data-cell')
        ? parseInt(el.dataset.colIdx, 10)
        : null;
    });
    expect(focused).toBe(0);
  });

  test('Tab at last cell in row in select mode stays at last cell', async ({ page }) => {
    const lastCell = page.locator('.subwindow table tbody tr:first-child td.data-cell').last();
    await lastCell.click();

    const colBefore = await page.evaluate(() => {
      const el = document.activeElement;
      return el && el.classList.contains('data-cell')
        ? parseInt(el.dataset.colIdx, 10)
        : null;
    });

    await page.keyboard.press('Tab');
    await page.waitForTimeout(200);

    // Should stay at the same cell (no crash, no wrap)
    const colAfter = await page.evaluate(() => {
      const el = document.activeElement;
      return el && el.classList.contains('data-cell')
        ? parseInt(el.dataset.colIdx, 10)
        : null;
    });
    expect(colAfter).toBe(colBefore);
  });

  test('Enter in select mode on last row stays on last row', async ({ page }) => {
    const lastRowCell = page.locator('.subwindow table tbody tr:last-child td.data-cell').first();
    await lastRowCell.click();

    const rowBefore = await page.evaluate(() => {
      const el = document.activeElement;
      return el && el.classList.contains('data-cell')
        ? parseInt(el.parentElement.dataset.displayIdx, 10)
        : null;
    });

    await page.keyboard.press('Enter');
    await page.waitForTimeout(200);

    // Should still be on the same row
    const rowAfter = await page.evaluate(() => {
      const el = document.activeElement;
      return el && el.classList.contains('data-cell')
        ? parseInt(el.parentElement.dataset.displayIdx, 10)
        : null;
    });
    expect(rowAfter).toBe(rowBefore);
  });

  test('multiple Tab then Enter navigates correctly across rows', async ({ page }) => {
    // Enter edit on column 0, row 1
    const cell = page.locator('.subwindow table tbody tr:first-child td.data-cell').first();
    await cell.click();
    await page.keyboard.press('F2');

    // Tab to column 1
    await page.keyboard.press('Tab');
    await page.waitForTimeout(100);

    // Enter should go to column 0 on row 2
    await page.keyboard.press('Enter');
    await page.waitForTimeout(200);

    const targetCell = page.locator('.subwindow table tbody tr:nth-child(2) td.data-cell').first();
    await expect(targetCell).toHaveAttribute('contenteditable', 'true');

    // Tab to column 1 again
    await page.keyboard.press('Tab');
    await page.waitForTimeout(100);

    // Enter should go to column 0 on row 3 (start column is still 0)
    await page.keyboard.press('Enter');
    await page.waitForTimeout(200);

    const thirdRowCell = page.locator('.subwindow table tbody tr:nth-child(3) td.data-cell').first();
    await expect(thirdRowCell).toHaveAttribute('contenteditable', 'true');
  });
});

// ---------------------------------------------------------------------------
// Undo/Redo: Column Rename
// ---------------------------------------------------------------------------

test.describe('Undo/Redo column rename', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('renameColumn pushes renameColumn entry to undo stack', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.renameColumn('sample1', 'name', 'full_name', win);
    });
    await page.waitForTimeout(300);

    const undoType = await page.evaluate(() => {
      const t = app._test.tables['sample1'];
      return t._undoStack && t._undoStack.length > 0
        ? t._undoStack[t._undoStack.length - 1].type
        : null;
    });
    expect(undoType).toBe('renameColumn');
  });

  test('renameColumn changes the column name', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.renameColumn('sample1', 'name', 'full_name', win);
    });
    await page.waitForTimeout(300);

    const data = await getTableData(page, 'sample1');
    expect(data.columns).toContain('full_name');
    expect(data.columns).not.toContain('name');
    expect(data.rows[0].full_name).toBe('Alice Johnson');
  });

  test('undo renameColumn restores the original column name', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.renameColumn('sample1', 'name', 'full_name', win);
    });
    await page.waitForTimeout(300);

    // Undo
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.undoTable('sample1', win);
    });
    await page.waitForTimeout(500);

    const data = await getTableData(page, 'sample1');
    expect(data.columns).toContain('name');
    expect(data.columns).not.toContain('full_name');
    expect(data.rows[0].name).toBe('Alice Johnson');
  });

  test('redo renameColumn re-applies the rename', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.renameColumn('sample1', 'name', 'full_name', win);
    });
    await page.waitForTimeout(300);

    // Undo
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.undoTable('sample1', win);
    });
    await page.waitForTimeout(500);

    // Redo
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.redoTable('sample1', win);
    });
    await page.waitForTimeout(500);

    const data = await getTableData(page, 'sample1');
    expect(data.columns).toContain('full_name');
    expect(data.columns).not.toContain('name');
    expect(data.rows[0].full_name).toBe('Alice Johnson');
  });

  test('undo renameColumn restores sort state', async ({ page }) => {
    // Set up a sort on the column, then rename, then undo
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      win.sortCols = [{ col: 'name', dir: 'asc' }];
    });

    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.renameColumn('sample1', 'name', 'full_name', win);
    });
    await page.waitForTimeout(300);

    // Verify sort was migrated to new name
    let sortCol = await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      return win.sortCols.length > 0 ? win.sortCols[0].col : null;
    });
    expect(sortCol).toBe('full_name');

    // Undo
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.undoTable('sample1', win);
    });
    await page.waitForTimeout(500);

    // Sort should be restored to original column name
    sortCol = await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      return win.sortCols.length > 0 ? win.sortCols[0].col : null;
    });
    expect(sortCol).toBe('name');
  });

  test('undo renameColumn restores column filter state', async ({ page }) => {
    // Set up a filter on the column, then rename, then undo
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      win.columnFilters['name'] = new Set(['Alice Johnson']);
    });

    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.renameColumn('sample1', 'name', 'full_name', win);
    });
    await page.waitForTimeout(300);

    // Undo
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.undoTable('sample1', win);
    });
    await page.waitForTimeout(500);

    const hasOrigFilter = await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      return 'name' in win.columnFilters;
    });
    expect(hasOrigFilter).toBe(true);
  });

  test('Edit menu shows "Undo Rename Column" after rename', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.renameColumn('sample1', 'name', 'full_name', win);
    });
    await page.waitForTimeout(300);

    // Focus a cell so getActiveDataWindow() works
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.waitForTimeout(100);

    await page.locator('#menu-edit .menu-label').click();
    await page.waitForTimeout(100);

    const undoLabel = await page.locator('#btn-undo').evaluate(el => el.firstChild.textContent);
    expect(undoLabel).toContain('Undo Rename Column');
    await page.keyboard.press('Escape');
  });

  test('Edit menu shows "Redo Rename Column" after undo', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.renameColumn('sample1', 'name', 'full_name', win);
    });
    await page.waitForTimeout(300);

    // Undo
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.undoTable('sample1', win);
    });
    await page.waitForTimeout(500);

    // Focus a cell
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click({ force: true });
    await page.waitForTimeout(100);

    await page.locator('#menu-edit .menu-label').click();
    await page.waitForTimeout(100);

    const redoLabel = await page.locator('#btn-redo').evaluate(el => el.firstChild.textContent);
    expect(redoLabel).toContain('Redo Rename Column');
    await page.keyboard.press('Escape');
  });

  test('undo renameColumn preserves data integrity across all rows', async ({ page }) => {
    const before = await getTableData(page, 'sample1');
    const originalNames = before.rows.map(r => r.name);

    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.renameColumn('sample1', 'name', 'full_name', win);
    });
    await page.waitForTimeout(300);

    // Undo
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.undoTable('sample1', win);
    });
    await page.waitForTimeout(500);

    const after = await getTableData(page, 'sample1');
    const restoredNames = after.rows.map(r => r.name);
    expect(restoredNames).toEqual(originalNames);
  });
});

// ---------------------------------------------------------------------------
// Undo/Redo: Column Reorder
// ---------------------------------------------------------------------------

test.describe('Undo/Redo column reorder', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('column reorder via drag pushes reorderColumn entry to undo stack', async ({ page }) => {
    // Drag column header to reorder
    const fromTh = page.locator('.subwindow table thead th[data-col-idx="0"]');
    const toTh = page.locator('.subwindow table thead th[data-col-idx="1"]');
    const fromBox = await fromTh.boundingBox();
    const toBox = await toTh.boundingBox();

    await page.mouse.move(fromBox.x + fromBox.width / 2, fromBox.y + fromBox.height / 2);
    await page.mouse.down();
    // Move past threshold
    await page.mouse.move(fromBox.x + fromBox.width / 2 + 10, fromBox.y + fromBox.height / 2, { steps: 3 });
    await page.mouse.move(toBox.x + toBox.width * 0.75, toBox.y + toBox.height / 2, { steps: 10 });
    await page.mouse.up();
    await page.waitForTimeout(300);

    const undoType = await page.evaluate(() => {
      const t = app._test.tables['sample1'];
      return t._undoStack && t._undoStack.length > 0
        ? t._undoStack[t._undoStack.length - 1].type
        : null;
    });
    expect(undoType).toBe('reorderColumn');
  });

  test('undo reorderColumn restores original column order', async ({ page }) => {
    const before = await getTableData(page, 'sample1');
    const originalOrder = [...before.columns];

    // Drag column 0 to after column 1
    const fromTh = page.locator('.subwindow table thead th[data-col-idx="0"]');
    const toTh = page.locator('.subwindow table thead th[data-col-idx="1"]');
    const fromBox = await fromTh.boundingBox();
    const toBox = await toTh.boundingBox();

    await page.mouse.move(fromBox.x + fromBox.width / 2, fromBox.y + fromBox.height / 2);
    await page.mouse.down();
    await page.mouse.move(fromBox.x + fromBox.width / 2 + 10, fromBox.y + fromBox.height / 2, { steps: 3 });
    await page.mouse.move(toBox.x + toBox.width * 0.75, toBox.y + toBox.height / 2, { steps: 10 });
    await page.mouse.up();
    await page.waitForTimeout(300);

    let data = await getTableData(page, 'sample1');
    // Columns should have changed
    expect(data.columns).not.toEqual(originalOrder);

    // Undo
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.undoTable('sample1', win);
    });
    await page.waitForTimeout(500);

    data = await getTableData(page, 'sample1');
    expect(data.columns).toEqual(originalOrder);
  });

  test('redo reorderColumn re-applies the reorder', async ({ page }) => {
    const before = await getTableData(page, 'sample1');

    // Drag column 0 to after column 1
    const fromTh = page.locator('.subwindow table thead th[data-col-idx="0"]');
    const toTh = page.locator('.subwindow table thead th[data-col-idx="1"]');
    const fromBox = await fromTh.boundingBox();
    const toBox = await toTh.boundingBox();

    await page.mouse.move(fromBox.x + fromBox.width / 2, fromBox.y + fromBox.height / 2);
    await page.mouse.down();
    await page.mouse.move(fromBox.x + fromBox.width / 2 + 10, fromBox.y + fromBox.height / 2, { steps: 3 });
    await page.mouse.move(toBox.x + toBox.width * 0.75, toBox.y + toBox.height / 2, { steps: 10 });
    await page.mouse.up();
    await page.waitForTimeout(300);

    const reordered = await getTableData(page, 'sample1');
    const reorderedOrder = [...reordered.columns];

    // Undo
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.undoTable('sample1', win);
    });
    await page.waitForTimeout(500);

    // Redo
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.redoTable('sample1', win);
    });
    await page.waitForTimeout(500);

    const data = await getTableData(page, 'sample1');
    expect(data.columns).toEqual(reorderedOrder);
  });

  test('undo reorderColumn restores column widths', async ({ page }) => {
    const widthsBefore = await page.evaluate(() => {
      const win = app._test.windows[0];
      return win.colWidths ? [...win.colWidths] : null;
    });

    // Drag column 0 to after column 1
    const fromTh = page.locator('.subwindow table thead th[data-col-idx="0"]');
    const toTh = page.locator('.subwindow table thead th[data-col-idx="1"]');
    const fromBox = await fromTh.boundingBox();
    const toBox = await toTh.boundingBox();

    await page.mouse.move(fromBox.x + fromBox.width / 2, fromBox.y + fromBox.height / 2);
    await page.mouse.down();
    await page.mouse.move(fromBox.x + fromBox.width / 2 + 10, fromBox.y + fromBox.height / 2, { steps: 3 });
    await page.mouse.move(toBox.x + toBox.width * 0.75, toBox.y + toBox.height / 2, { steps: 10 });
    await page.mouse.up();
    await page.waitForTimeout(300);

    // Undo
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.undoTable('sample1', win);
    });
    await page.waitForTimeout(500);

    const widthsAfter = await page.evaluate(() => {
      const win = app._test.windows[0];
      return win.colWidths ? [...win.colWidths] : null;
    });

    if (widthsBefore && widthsAfter) {
      expect(widthsAfter).toEqual(widthsBefore);
    }
  });

  test('Edit menu shows "Undo Reorder Column" after reorder', async ({ page }) => {
    // Drag column 0 to after column 1
    const fromTh = page.locator('.subwindow table thead th[data-col-idx="0"]');
    const toTh = page.locator('.subwindow table thead th[data-col-idx="1"]');
    const fromBox = await fromTh.boundingBox();
    const toBox = await toTh.boundingBox();

    await page.mouse.move(fromBox.x + fromBox.width / 2, fromBox.y + fromBox.height / 2);
    await page.mouse.down();
    await page.mouse.move(fromBox.x + fromBox.width / 2 + 10, fromBox.y + fromBox.height / 2, { steps: 3 });
    await page.mouse.move(toBox.x + toBox.width * 0.75, toBox.y + toBox.height / 2, { steps: 10 });
    await page.mouse.up();
    await page.waitForTimeout(300);

    // Focus a cell so getActiveDataWindow() works
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.waitForTimeout(100);

    await page.locator('#menu-edit .menu-label').click();
    await page.waitForTimeout(100);

    const undoLabel = await page.locator('#btn-undo').evaluate(el => el.firstChild.textContent);
    expect(undoLabel).toContain('Undo Reorder Column');
    await page.keyboard.press('Escape');
  });

  test('Edit menu shows "Redo Reorder Column" after undo', async ({ page }) => {
    // Reorder
    const fromTh = page.locator('.subwindow table thead th[data-col-idx="0"]');
    const toTh = page.locator('.subwindow table thead th[data-col-idx="1"]');
    const fromBox = await fromTh.boundingBox();
    const toBox = await toTh.boundingBox();

    await page.mouse.move(fromBox.x + fromBox.width / 2, fromBox.y + fromBox.height / 2);
    await page.mouse.down();
    await page.mouse.move(fromBox.x + fromBox.width / 2 + 10, fromBox.y + fromBox.height / 2, { steps: 3 });
    await page.mouse.move(toBox.x + toBox.width * 0.75, toBox.y + toBox.height / 2, { steps: 10 });
    await page.mouse.up();
    await page.waitForTimeout(300);

    // Undo
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.undoTable('sample1', win);
    });
    await page.waitForTimeout(500);

    // Focus a cell
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click({ force: true });
    await page.waitForTimeout(100);

    await page.locator('#menu-edit .menu-label').click();
    await page.waitForTimeout(100);

    const redoLabel = await page.locator('#btn-redo').evaluate(el => el.firstChild.textContent);
    expect(redoLabel).toContain('Redo Reorder Column');
    await page.keyboard.press('Escape');
  });
});

// ---------------------------------------------------------------------------
// Undo/Redo: Column Resize
// ---------------------------------------------------------------------------

test.describe('Undo/Redo column resize', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('drag resize pushes resizeColumn entry to undo stack', async ({ page }) => {
    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 102, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const undoType = await page.evaluate(() => {
      const t = app._test.tables['sample1'];
      return t._undoStack && t._undoStack.length > 0
        ? t._undoStack[t._undoStack.length - 1].type
        : null;
    });
    expect(undoType).toBe('resizeColumn');
  });

  test('double-click auto-fit pushes resizeColumn entry to undo stack', async ({ page }) => {
    // Widen column first
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.colWidths[0] = 400;
      win._colgroup.children[1].style.width = '400px';
      app._test.updateTableWidth(win);
    });

    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
    const box = await handle.boundingBox();
    await page.mouse.dblclick(box.x + 2, box.y + box.height / 2);
    await page.waitForTimeout(300);

    const undoType = await page.evaluate(() => {
      const t = app._test.tables['sample1'];
      return t._undoStack && t._undoStack.length > 0
        ? t._undoStack[t._undoStack.length - 1].type
        : null;
    });
    expect(undoType).toBe('resizeColumn');
  });

  test('undo resizeColumn restores old column widths', async ({ page }) => {
    const widthsBefore = await page.evaluate(() => {
      const win = app._test.windows[0];
      return [...win.colWidths];
    });

    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 102, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    // Undo
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.undoTable('sample1', win);
    });
    await page.waitForTimeout(500);

    const widthsAfter = await page.evaluate(() => {
      const win = app._test.windows[0];
      return [...win.colWidths];
    });

    expect(widthsAfter).toEqual(widthsBefore);
  });

  test('redo resizeColumn re-applies the new widths', async ({ page }) => {
    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 102, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const widthsResized = await page.evaluate(() => {
      const win = app._test.windows[0];
      return [...win.colWidths];
    });

    // Undo
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.undoTable('sample1', win);
    });
    await page.waitForTimeout(500);

    // Redo
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.redoTable('sample1', win);
    });
    await page.waitForTimeout(500);

    const widthsRedo = await page.evaluate(() => {
      const win = app._test.windows[0];
      return [...win.colWidths];
    });

    expect(widthsRedo).toEqual(widthsResized);
  });

  test('resize undo is not pushed when widths do not change', async ({ page }) => {
    // Clear any existing undo entries
    await page.evaluate(() => {
      const t = app._test.tables['sample1'];
      t._undoStack = [];
      t._redoStack = [];
    });

    // Mousedown and mouseup at the same position (no drag)
    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.up();
    await page.waitForTimeout(200);

    const stackLen = await page.evaluate(() => {
      const t = app._test.tables['sample1'];
      return t._undoStack ? t._undoStack.length : 0;
    });
    expect(stackLen).toBe(0);
  });

  test('Edit menu shows "Undo Resize Column" after resize', async ({ page }) => {
    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 102, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    // Focus a cell so getActiveDataWindow() works
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.waitForTimeout(100);

    await page.locator('#menu-edit .menu-label').click();
    await page.waitForTimeout(100);

    const undoLabel = await page.locator('#btn-undo').evaluate(el => el.firstChild.textContent);
    expect(undoLabel).toContain('Undo Resize Column');
    await page.keyboard.press('Escape');
  });

  test('Edit menu shows "Redo Resize Column" after undo', async ({ page }) => {
    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 102, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    // Undo
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.undoTable('sample1', win);
    });
    await page.waitForTimeout(500);

    // Focus a cell
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click({ force: true });
    await page.waitForTimeout(100);

    await page.locator('#menu-edit .menu-label').click();
    await page.waitForTimeout(100);

    const redoLabel = await page.locator('#btn-redo').evaluate(el => el.firstChild.textContent);
    expect(redoLabel).toContain('Redo Resize Column');
    await page.keyboard.press('Escape');
  });

  test('undo/redo resize round-trips correctly', async ({ page }) => {
    const originalWidths = await page.evaluate(() => {
      const win = app._test.windows[0];
      return [...win.colWidths];
    });

    // Resize
    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 102, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    // Undo
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.undoTable('sample1', win);
    });
    await page.waitForTimeout(500);

    const afterUndo = await page.evaluate(() => {
      const win = app._test.windows[0];
      return [...win.colWidths];
    });
    expect(afterUndo).toEqual(originalWidths);

    // Redo
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.redoTable('sample1', win);
    });
    await page.waitForTimeout(500);

    const afterRedo = await page.evaluate(() => {
      const win = app._test.windows[0];
      return [...win.colWidths];
    });
    expect(afterRedo[0]).toBeGreaterThan(originalWidths[0] + 50);

    // Undo again
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.undoTable('sample1', win);
    });
    await page.waitForTimeout(500);

    const afterSecondUndo = await page.evaluate(() => {
      const win = app._test.windows[0];
      return [...win.colWidths];
    });
    expect(afterSecondUndo).toEqual(originalWidths);
  });
});

// ---------------------------------------------------------------------------
// Drag resize: no snap on click-without-drag
// ---------------------------------------------------------------------------

test.describe('Drag resize no-snap on click without drag', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('mousedown+mouseup without mousemove does not change other selected columns', async ({ page }) => {
    // Select all cells so all columns are in the selection
    await page.evaluate(() => {
      const win = app._test.windows[0];
      app._test.selectAllCells(win);
    });
    await page.waitForTimeout(100);

    const widthsBefore = await page.evaluate(() => {
      const win = app._test.windows[0];
      return [...win.colWidths];
    });

    // Click (mousedown + mouseup) on the first column's resize handle without moving
    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.up();
    await page.waitForTimeout(200);

    const widthsAfter = await page.evaluate(() => {
      const win = app._test.windows[0];
      return [...win.colWidths];
    });

    // All column widths should remain unchanged
    expect(widthsAfter).toEqual(widthsBefore);
  });

  test('mousedown+mouseup without drag does not push undo entry', async ({ page }) => {
    // Clear undo stack
    await page.evaluate(() => {
      const t = app._test.tables['sample1'];
      t._undoStack = [];
      t._redoStack = [];
    });

    // Select all cells
    await page.evaluate(() => {
      const win = app._test.windows[0];
      app._test.selectAllCells(win);
    });
    await page.waitForTimeout(100);

    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.up();
    await page.waitForTimeout(200);

    const stackLen = await page.evaluate(() => {
      const t = app._test.tables['sample1'];
      return t._undoStack ? t._undoStack.length : 0;
    });
    expect(stackLen).toBe(0);
  });

  test('actual drag resize with selection still snaps all selected columns', async ({ page }) => {
    // Select all cells
    await page.evaluate(() => {
      const win = app._test.windows[0];
      app._test.selectAllCells(win);
    });
    await page.waitForTimeout(100);

    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 102, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const widthsAfter = await page.evaluate(() => {
      const win = app._test.windows[0];
      return [...win.colWidths];
    });

    // All columns should have the same width
    expect(widthsAfter[0]).toBe(widthsAfter[1]);
    expect(widthsAfter[1]).toBe(widthsAfter[2]);
  });

  test('click-without-drag on unselected column handle does not affect widths', async ({ page }) => {
    // Select only column 1
    await page.evaluate(() => {
      const win = app._test.windows[0];
      const cols = win._columns;
      const rows = win._displayRows;
      win.anchorCell = { rownum: rows[0]._rownum, col: cols[1] };
      win.selectedCells = new Set([`${rows[0]._rownum}:${cols[1]}`]);
    });
    await page.waitForTimeout(100);

    const widthsBefore = await page.evaluate(() => {
      const win = app._test.windows[0];
      return [...win.colWidths];
    });

    // Click the first column's resize handle (column 0 is not in selection)
    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.up();
    await page.waitForTimeout(200);

    const widthsAfter = await page.evaluate(() => {
      const win = app._test.windows[0];
      return [...win.colWidths];
    });

    expect(widthsAfter).toEqual(widthsBefore);
  });
});

// ---------------------------------------------------------------------------
// Global Ctrl+Z/Ctrl+Shift+Z without cell focus
// ---------------------------------------------------------------------------

test.describe('Global undo/redo without cell focus', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('Ctrl+Z from title bar undoes last operation', async ({ page }) => {
    // Make an edit via renameColumn (no clipboard permissions needed)
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.renameColumn('sample1', 'name', 'renamed_col', win);
    });
    await page.waitForTimeout(300);

    let data = await getTableData(page, 'sample1');
    expect(data.columns).toContain('renamed_col');

    // Click the title bar (not a cell, not an input)
    await page.locator('.subwindow .win-title').first().click();
    await page.waitForTimeout(100);

    // Ctrl+Z from the title bar should trigger the global handler
    await page.keyboard.press('Control+z');
    await page.waitForTimeout(500);

    data = await getTableData(page, 'sample1');
    expect(data.columns).toContain('name');
    expect(data.columns).not.toContain('renamed_col');
  });

  test('Ctrl+Shift+Z from title bar redoes last operation', async ({ page }) => {
    // Make an edit
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.renameColumn('sample1', 'name', 'renamed_col', win);
    });
    await page.waitForTimeout(300);

    // Undo via evaluate
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.undoTable('sample1', win);
    });
    await page.waitForTimeout(500);

    let data = await getTableData(page, 'sample1');
    expect(data.columns).toContain('name');

    // Click title bar and redo
    await page.locator('.subwindow .win-title').first().click();
    await page.waitForTimeout(100);

    await page.keyboard.press('Control+Shift+z');
    await page.waitForTimeout(500);

    data = await getTableData(page, 'sample1');
    expect(data.columns).toContain('renamed_col');
    expect(data.columns).not.toContain('name');
  });

  test('global Ctrl+Z does not fire when focus is in an input', async ({ page }) => {
    // Make an edit via rename to create an undo entry
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.renameColumn('sample1', 'name', 'full_name', win);
    });
    await page.waitForTimeout(300);

    // Focus the SQL input
    await page.locator('#sql-input').click();
    await page.locator('#sql-input').fill('test');
    await page.waitForTimeout(100);

    // Ctrl+Z should be handled by the input, not the global handler
    await page.keyboard.press('Control+z');
    await page.waitForTimeout(300);

    // The rename should NOT have been undone
    const data = await getTableData(page, 'sample1');
    expect(data.columns).toContain('full_name');
    expect(data.columns).not.toContain('name');
  });

  test('global Ctrl+Z does not fire when focus is in a contenteditable cell', async ({ page }) => {
    // Create an undo entry first
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.renameColumn('sample1', 'name', 'full_name', win);
    });
    await page.waitForTimeout(300);

    // Enter edit mode on a cell
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.keyboard.press('F2');
    await expect(cell).toHaveAttribute('contenteditable', 'true');

    // Ctrl+Z in edit mode should be handled by contenteditable, not global handler
    await page.keyboard.press('Control+z');
    await page.waitForTimeout(300);

    // The rename should NOT have been undone via the global handler
    const data = await getTableData(page, 'sample1');
    expect(data.columns).toContain('full_name');
  });

  test('global undo works for structural operations from title bar', async ({ page }) => {
    const before = await getTableData(page, 'sample1');

    // Insert a row
    const rowNum = page.locator('.subwindow table tbody td.row-num').first();
    await rowNum.click({ button: 'right' });
    await page.locator('.context-menu button', { hasText: 'Insert Row Below' }).click();
    await page.waitForTimeout(300);

    let data = await getTableData(page, 'sample1');
    expect(data.rows.length).toBe(before.rows.length + 1);

    // Click title bar and undo
    await page.locator('.subwindow .win-title').first().click();
    await page.waitForTimeout(100);

    await page.keyboard.press('Control+z');
    await page.waitForTimeout(500);

    data = await getTableData(page, 'sample1');
    expect(data.rows.length).toBe(before.rows.length);
  });

  test('global undo works for resize operations from title bar', async ({ page }) => {
    const widthsBefore = await page.evaluate(() => {
      const win = app._test.windows[0];
      return [...win.colWidths];
    });

    // Resize a column
    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 102, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    // Click title bar and undo
    await page.locator('.subwindow .win-title').first().click();
    await page.waitForTimeout(100);

    await page.keyboard.press('Control+z');
    await page.waitForTimeout(500);

    const widthsAfter = await page.evaluate(() => {
      const win = app._test.windows[0];
      return [...win.colWidths];
    });

    expect(widthsAfter).toEqual(widthsBefore);
  });

  test('global Ctrl+Z works from window body (non-cell, non-input area)', async ({ page }) => {
    // Make an edit via renameColumn (no clipboard permissions needed)
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.renameColumn('sample1', 'name', 'body_undo_col', win);
    });
    await page.waitForTimeout(300);

    let data = await getTableData(page, 'sample1');
    expect(data.columns).toContain('body_undo_col');

    // Click somewhere on the window that is not a cell or input
    // The status bar area should work
    await page.locator('.subwindow .win-statusbar').first().click();
    await page.waitForTimeout(100);

    await page.keyboard.press('Control+z');
    await page.waitForTimeout(500);

    data = await getTableData(page, 'sample1');
    expect(data.columns).toContain('name');
    expect(data.columns).not.toContain('body_undo_col');
  });
});
