const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, getTableData } = require('../helpers');

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

test.describe('Copy, Paste, Undo, Redo', () => {
  test.beforeEach(async ({ page, context }) => {
    await context.grantPermissions(['clipboard-read', 'clipboard-write']);
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('Ctrl+C copies a single cell to clipboard as plain text', async ({ page }) => {
    const firstCell = page.locator('.subwindow table tbody td.data-cell').first();
    await firstCell.click();
    await page.keyboard.press('Control+c');
    const clip = await page.evaluate(() => navigator.clipboard.readText());
    expect(clip).toBe('Alice Johnson');
  });

  test('Ctrl+C copies a rectangular selection as TSV', async ({ page }) => {
    const firstCell = page.locator('.subwindow table tbody td.data-cell').first();
    await firstCell.click();
    await page.keyboard.press('Shift+ArrowRight');
    await page.keyboard.press('Shift+ArrowDown');
    await page.waitForTimeout(100);
    await page.keyboard.press('Control+c');
    const clip = await page.evaluate(() => navigator.clipboard.readText());
    expect(clip).toBe('Alice Johnson\talice@example.com\nBob Smith\tbob@example.com');
  });

  test('Ctrl+V pastes a single value into the focused cell', async ({ page }) => {
    await page.evaluate(() => navigator.clipboard.writeText('PASTED'));
    const firstCell = page.locator('.subwindow table tbody td.data-cell').first();
    await firstCell.click();
    await page.keyboard.press('Control+v');
    await page.waitForTimeout(300);
    const data = await getTableData(page, 'sample1');
    expect(data.rows[0].name).toBe('PASTED');
    expect(data.modified).toBe(true);
  });

  test('Ctrl+V pastes multi-cell TSV starting from anchor', async ({ page }) => {
    const tsv = 'X1\tX2\nY1\tY2';
    await page.evaluate((t) => navigator.clipboard.writeText(t), tsv);
    const firstCell = page.locator('.subwindow table tbody td.data-cell').first();
    await firstCell.click();
    await page.keyboard.press('Control+v');
    await page.waitForTimeout(300);
    const data = await getTableData(page, 'sample1');
    expect(data.rows[0].name).toBe('X1');
    expect(data.rows[0].email).toBe('X2');
    expect(data.rows[1].name).toBe('Y1');
    expect(data.rows[1].email).toBe('Y2');
  });

  test('paste clips to table bounds', async ({ page }) => {
    const tsv = 'A\tB\tC\tD\tE';
    await page.evaluate((t) => navigator.clipboard.writeText(t), tsv);
    // Click last column of first row
    const cells = page.locator('.subwindow table tbody tr:first-child td.data-cell');
    const lastCell = cells.last();
    await lastCell.click();
    await page.keyboard.press('Control+v');
    await page.waitForTimeout(300);
    const data = await getTableData(page, 'sample1');
    // Only the last column should have changed (member_since is the 3rd and last column)
    expect(data.rows[0].member_since).toBe('A');
    // Other rows unchanged
    expect(data.rows[1].name).toBe('Bob Smith');
  });

  test('Ctrl+Z undoes a single cell edit', async ({ page }) => {
    const firstCell = page.locator('.subwindow table tbody td.data-cell').first();
    await firstCell.click();
    // Edit cell by directly setting content
    await page.evaluate(() => {
      const t = app._test.tables['sample1'];
      const row = t.rows[0];
      const oldVal = row.name;
      row.name = 'NewValue';
      if (!t._dirtyCells) t._dirtyCells = [];
      t._dirtyCells.push({ rownum: 1, col: 'name', value: 'NewValue' });
      if (!t._undoStack) t._undoStack = [];
      t._undoStack.push({ type: 'edit', changes: [{ rownum: 1, col: 'name', oldValue: oldVal, newValue: 'NewValue' }] });
      t._redoStack = [];
      t.modified = true;
    });
    await page.evaluate(() => app._test.rebuildTable(app._test.windows.find(w => w.tableName === 'sample1')));
    await page.waitForTimeout(100);
    let data = await getTableData(page, 'sample1');
    expect(data.rows[0].name).toBe('NewValue');

    // Click cell again then undo
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.keyboard.press('Control+z');
    await page.waitForTimeout(200);
    data = await getTableData(page, 'sample1');
    expect(data.rows[0].name).toBe('Alice Johnson');
  });

  test('Ctrl+Z undoes a paste as a single step', async ({ page }) => {
    const tsv = 'P1\tP2\nP3\tP4';
    await page.evaluate((t) => navigator.clipboard.writeText(t), tsv);
    const firstCell = page.locator('.subwindow table tbody td.data-cell').first();
    await firstCell.click();
    await page.keyboard.press('Control+v');
    await page.waitForTimeout(300);
    let data = await getTableData(page, 'sample1');
    expect(data.rows[0].name).toBe('P1');
    expect(data.rows[1].email).toBe('P4');

    // Undo reverts all 4 cells at once
    await firstCell.click();
    await page.keyboard.press('Control+z');
    await page.waitForTimeout(100);
    data = await getTableData(page, 'sample1');
    expect(data.rows[0].name).toBe('Alice Johnson');
    expect(data.rows[0].email).toBe('alice@example.com');
    expect(data.rows[1].name).toBe('Bob Smith');
    expect(data.rows[1].email).toBe('bob@example.com');
  });

  test('Ctrl+Shift+Z redoes an undone edit', async ({ page }) => {
    // Set up an edit with undo entry
    await page.evaluate(() => {
      const t = app._test.tables['sample1'];
      const row = t.rows[0];
      row.name = 'Edited';
      if (!t._dirtyCells) t._dirtyCells = [];
      t._dirtyCells.push({ rownum: 1, col: 'name', value: 'Edited' });
      if (!t._undoStack) t._undoStack = [];
      t._undoStack.push({ type: 'edit', changes: [{ rownum: 1, col: 'name', oldValue: 'Alice Johnson', newValue: 'Edited' }] });
      t._redoStack = [];
      t.modified = true;
    });
    await page.evaluate(() => app._test.rebuildTable(app._test.windows.find(w => w.tableName === 'sample1')));
    await page.waitForTimeout(100);

    // Undo
    const firstCell = page.locator('.subwindow table tbody td.data-cell').first();
    await firstCell.click();
    await page.keyboard.press('Control+z');
    await page.waitForTimeout(200);
    let data = await getTableData(page, 'sample1');
    expect(data.rows[0].name).toBe('Alice Johnson');

    // Redo — need to click cell again since undo re-renders
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.keyboard.press('Control+Shift+z');
    await page.waitForTimeout(200);
    data = await getTableData(page, 'sample1');
    expect(data.rows[0].name).toBe('Edited');
  });

  test('new edit clears the redo stack', async ({ page }) => {
    // Set up V1 edit
    await page.evaluate(() => {
      const t = app._test.tables['sample1'];
      t.rows[0].name = 'V1';
      if (!t._undoStack) t._undoStack = [];
      t._undoStack.push({ type: 'edit', changes: [{ rownum: 1, col: 'name', oldValue: 'Alice Johnson', newValue: 'V1' }] });
      t._redoStack = [];
    });
    await page.evaluate(() => app._test.rebuildTable(app._test.windows.find(w => w.tableName === 'sample1')));

    // Undo V1
    const firstCell = page.locator('.subwindow table tbody td.data-cell').first();
    await firstCell.click();
    await page.keyboard.press('Control+z');
    await page.waitForTimeout(200);

    // Make a different edit (V2) via paste to avoid contenteditable issues
    await page.evaluate(() => navigator.clipboard.writeText('V2'));
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.keyboard.press('Control+v');
    await page.waitForTimeout(300);
    let data = await getTableData(page, 'sample1');
    expect(data.rows[0].name).toBe('V2');

    // Redo should do nothing (stack cleared by new paste)
    const cell2 = page.locator('.subwindow table tbody td.data-cell').first();
    await cell2.click();
    await page.keyboard.press('Control+Shift+z');
    await page.waitForTimeout(200);
    data = await getTableData(page, 'sample1');
    expect(data.rows[0].name).toBe('V2');
  });

  test('copy and paste round-trips correctly', async ({ page }) => {
    // Select a 2x2 block and copy
    const firstCell = page.locator('.subwindow table tbody td.data-cell').first();
    await firstCell.click();
    await page.keyboard.press('Shift+ArrowRight');
    await page.keyboard.press('Shift+ArrowDown');
    await page.waitForTimeout(100);
    await page.keyboard.press('Control+c');

    // Move to row 5 and paste
    const targetCell = page.locator('.subwindow table tbody tr:nth-child(5) td.data-cell').first();
    await targetCell.click();
    await page.keyboard.press('Control+v');
    await page.waitForTimeout(300);

    const data = await getTableData(page, 'sample1');
    expect(data.rows[4].name).toBe('Alice Johnson');
    expect(data.rows[4].email).toBe('alice@example.com');
    expect(data.rows[5].name).toBe('Bob Smith');
    expect(data.rows[5].email).toBe('bob@example.com');
  });

  test('Ctrl+A selects all cells', async ({ page }) => {
    const firstCell = page.locator('.subwindow table tbody td.data-cell').first();
    await firstCell.click();
    await page.keyboard.press('Control+a');
    await page.waitForTimeout(100);
    const sel = await getSelection(page, 'sample1');
    // 10 rows x 3 columns = 30 cells
    expect(sel.cells.length).toBe(30);
  });

  test('clicking # corner cell selects all cells', async ({ page }) => {
    const corner = page.locator('.subwindow table thead th.row-num-header');
    await corner.click();
    await page.waitForTimeout(100);
    const sel = await getSelection(page, 'sample1');
    expect(sel.cells.length).toBe(30);
  });

  test('clicking a row number selects that row', async ({ page }) => {
    const rowNum = page.locator('.subwindow table tbody td.row-num').first();
    await rowNum.click();
    await page.waitForTimeout(100);
    const sel = await getSelection(page, 'sample1');
    // 1 row x 3 columns = 3 cells
    expect(sel.cells.length).toBe(3);
    expect(sel.cells).toEqual(['1:email', '1:member_since', '1:name']);
  });

  test('shift-clicking row numbers selects a range of rows', async ({ page }) => {
    const row1 = page.locator('.subwindow table tbody td.row-num').nth(0);
    const row3 = page.locator('.subwindow table tbody td.row-num').nth(2);
    await row1.click();
    await row3.click({ modifiers: ['Shift'] });
    await page.waitForTimeout(100);
    const sel = await getSelection(page, 'sample1');
    // 3 rows x 3 columns = 9 cells
    expect(sel.cells.length).toBe(9);
  });

  test('select-all copy includes header row', async ({ page }) => {
    const corner = page.locator('.subwindow table thead th.row-num-header');
    await corner.click();
    await page.waitForTimeout(100);
    await page.keyboard.press('Control+c');
    const clip = await page.evaluate(() => navigator.clipboard.readText());
    const lines = clip.split('\n');
    expect(lines[0]).toBe('name\temail\tmember_since');
    expect(lines[1].startsWith('Alice Johnson\t')).toBe(true);
    expect(lines.length).toBe(11); // 1 header + 10 data rows
  });

  test('row-select copy includes header row', async ({ page }) => {
    const rowNum = page.locator('.subwindow table tbody td.row-num').first();
    await rowNum.click();
    await page.waitForTimeout(100);
    await page.keyboard.press('Control+c');
    const clip = await page.evaluate(() => navigator.clipboard.readText());
    const lines = clip.split('\n');
    expect(lines[0]).toBe('name\temail\tmember_since');
    expect(lines[1].startsWith('Alice Johnson\t')).toBe(true);
    expect(lines.length).toBe(2); // 1 header + 1 data row
  });

  test('normal cell selection copy does not include header', async ({ page }) => {
    const firstCell = page.locator('.subwindow table tbody td.data-cell').first();
    await firstCell.click();
    await page.keyboard.press('Control+c');
    const clip = await page.evaluate(() => navigator.clipboard.readText());
    expect(clip).toBe('Alice Johnson');
  });

  test('Ctrl+X cuts selected cells (copies and clears)', async ({ page }) => {
    const firstCell = page.locator('.subwindow table tbody td.data-cell').first();
    await firstCell.click();
    await page.keyboard.press('Shift+ArrowRight');
    await page.waitForTimeout(100);
    await page.keyboard.press('Control+x');
    await page.waitForTimeout(200);
    const clip = await page.evaluate(() => navigator.clipboard.readText());
    expect(clip).toBe('Alice Johnson\talice@example.com');
    const data = await getTableData(page, 'sample1');
    expect(data.rows[0].name).toBe('');
    expect(data.rows[0].email).toBe('');
  });

  test('Ctrl+Z undoes a cut operation', async ({ page }) => {
    const firstCell = page.locator('.subwindow table tbody td.data-cell').first();
    await firstCell.click();
    await page.keyboard.press('Control+x');
    await page.waitForTimeout(200);
    let data = await getTableData(page, 'sample1');
    expect(data.rows[0].name).toBe('');
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.keyboard.press('Control+z');
    await page.waitForTimeout(200);
    data = await getTableData(page, 'sample1');
    expect(data.rows[0].name).toBe('Alice Johnson');
  });

  test('dragging row numbers selects a range of rows', async ({ page }) => {
    const row1 = page.locator('.subwindow table tbody td.row-num').nth(0);
    const row3 = page.locator('.subwindow table tbody td.row-num').nth(2);
    const box1 = await row1.boundingBox();
    const box3 = await row3.boundingBox();
    await page.mouse.move(box1.x + box1.width / 2, box1.y + box1.height / 2);
    await page.mouse.down();
    await page.mouse.move(box3.x + box3.width / 2, box3.y + box3.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(100);
    const sel = await getSelection(page, 'sample1');
    expect(sel.cells.length).toBe(9); // 3 rows x 3 cols
  });

  test('Edit menu shows Undo/Redo with action labels', async ({ page }) => {
    // Make an edit via paste
    await page.evaluate(() => navigator.clipboard.writeText('TEST'));
    const firstCell = page.locator('.subwindow table tbody td.data-cell').first();
    await firstCell.click();
    await page.keyboard.press('Control+v');
    await page.waitForTimeout(300);
    // Open Edit menu
    await page.locator('#menu-edit .menu-label').click();
    await page.waitForTimeout(100);
    const undoBtn = page.locator('#btn-undo');
    await expect(undoBtn).toBeEnabled();
    const undoText = await undoBtn.innerText();
    expect(undoText).toContain('Undo Paste');
    const redoBtn = page.locator('#btn-redo');
    await expect(redoBtn).toBeDisabled();
    // Close menu
    await page.keyboard.press('Escape');
  });

  test('Edit menu Select All button works', async ({ page }) => {
    await page.locator('#menu-edit .menu-label').click();
    await page.waitForTimeout(100);
    const selAllBtn = page.locator('#btn-select-all');
    await expect(selAllBtn).toBeEnabled();
    await selAllBtn.click();
    await page.waitForTimeout(100);
    const sel = await getSelection(page, 'sample1');
    expect(sel.cells.length).toBe(30);
  });

  test('Ctrl+A from global handler works without cell focus', async ({ page }) => {
    // Click somewhere that's not a data cell (e.g., the window title bar)
    await page.locator('.subwindow .win-title').first().click();
    await page.waitForTimeout(100);
    await page.keyboard.press('Control+a');
    await page.waitForTimeout(100);
    const sel = await getSelection(page, 'sample1');
    expect(sel.cells.length).toBe(30);
  });

  test('Escape deselects all cells via keyboard', async ({ page }) => {
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.keyboard.press('Control+a');
    await page.waitForTimeout(100);
    let sel = await getSelection(page, 'sample1');
    expect(sel.cells.length).toBe(30);
    // Escape deselects all
    await page.keyboard.press('Escape');
    await page.waitForTimeout(100);
    sel = await getSelection(page, 'sample1');
    expect(sel.cells.length).toBe(0);
    expect(sel.anchor).toBeNull();
  });

  test('Edit menu Select None button deselects all cells', async ({ page }) => {
    // Select all first
    await page.locator('#menu-edit .menu-label').click();
    await page.waitForTimeout(100);
    await page.locator('#btn-select-all').click();
    await page.waitForTimeout(100);
    let sel = await getSelection(page, 'sample1');
    expect(sel.cells.length).toBe(30);
    // Click Select None
    await page.locator('#menu-edit .menu-label').click();
    await page.waitForTimeout(100);
    await page.locator('#btn-select-none').click();
    await page.waitForTimeout(100);
    sel = await getSelection(page, 'sample1');
    expect(sel.cells.length).toBe(0);
    expect(sel.anchor).toBeNull();
  });

  test('Select All toggles to Select None via corner cell click', async ({ page }) => {
    const corner = page.locator('.subwindow table thead th.row-num-header');
    await corner.click();
    await page.waitForTimeout(100);
    let sel = await getSelection(page, 'sample1');
    expect(sel.cells.length).toBe(30);
    // Second click deselects
    await corner.click();
    await page.waitForTimeout(100);
    sel = await getSelection(page, 'sample1');
    expect(sel.cells.length).toBe(0);
    expect(sel.anchor).toBeNull();
  });

  test('Edit menu Select None is disabled when no cells are selected', async ({ page }) => {
    await page.locator('#menu-edit .menu-label').click();
    await page.waitForTimeout(100);
    const isDisabled = await page.locator('#btn-select-none').evaluate(el => el.disabled);
    expect(isDisabled).toBe(true);
  });

  test('Edit menu Select None is enabled when cells are selected', async ({ page }) => {
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.waitForTimeout(100);
    await page.locator('#menu-edit .menu-label').click();
    await page.waitForTimeout(100);
    const isDisabled = await page.locator('#btn-select-none').evaluate(el => el.disabled);
    expect(isDisabled).toBe(false);
  });

  test('Select All always selects even when already all selected', async ({ page }) => {
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.keyboard.press('Control+a');
    await page.waitForTimeout(100);
    let sel = await getSelection(page, 'sample1');
    expect(sel.cells.length).toBe(30);
    // Pressing Ctrl+A again does not deselect
    await page.keyboard.press('Control+a');
    await page.waitForTimeout(100);
    sel = await getSelection(page, 'sample1');
    expect(sel.cells.length).toBe(30);
  });

  test('Select None clears _copyWithHeader flag', async ({ page }) => {
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.keyboard.press('Control+a');
    await page.waitForTimeout(100);
    // Verify _copyWithHeader is set
    let copyWithHeader = await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'sample1');
      return w._copyWithHeader;
    });
    expect(copyWithHeader).toBe(true);
    // Escape to select none
    await page.keyboard.press('Escape');
    await page.waitForTimeout(100);
    copyWithHeader = await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'sample1');
      return w._copyWithHeader;
    });
    expect(copyWithHeader).toBe(false);
  });

  test('Edit menu Select None deselects without cell focus', async ({ page }) => {
    // Select all from title bar (global handler)
    await page.locator('.subwindow .win-title').first().click();
    await page.waitForTimeout(100);
    await page.keyboard.press('Control+a');
    await page.waitForTimeout(100);
    let sel = await getSelection(page, 'sample1');
    expect(sel.cells.length).toBe(30);
    // Deselect via Edit menu
    await page.locator('#menu-edit .menu-label').click();
    await page.waitForTimeout(100);
    await page.locator('#btn-select-none').click();
    await page.waitForTimeout(100);
    sel = await getSelection(page, 'sample1');
    expect(sel.cells.length).toBe(0);
  });

  test('Escape in edit mode reverts edit instead of deselecting', async ({ page }) => {
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.waitForTimeout(50);
    const original = await cell.textContent();
    // Enter edit mode
    await page.keyboard.press('F2');
    await page.waitForTimeout(50);
    await page.keyboard.type('CHANGED');
    // Escape should revert the edit
    await page.keyboard.press('Escape');
    await page.waitForTimeout(100);
    const after = await cell.textContent();
    expect(after).toBe(original);
    // Cell should still be focused (selection not cleared)
    const sel = await getSelection(page, 'sample1');
    expect(sel.cells.length).toBeGreaterThan(0);
  });

  test('Select None shortcut shows Esc in Edit menu', async ({ page }) => {
    await page.locator('#menu-edit .menu-label').click();
    await page.waitForTimeout(100);
    const shortcut = await page.locator('#btn-select-none .shortcut').textContent();
    expect(shortcut).toBe('Esc');
  });

  test('Escape clears selectedCol and _copyWithHeader', async ({ page }) => {
    // Click a data cell first to ensure focus is on the table
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.waitForTimeout(50);
    // Select a column via programmatic call (sets selectedCol + _copyWithHeader)
    await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.selectColumns(w, 0, 0);
      w.selectedCol = w._columns[0];
    });
    await page.waitForTimeout(100);
    let state = await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'sample1');
      return { selectedCol: w.selectedCol, copyWithHeader: w._copyWithHeader, cells: w.selectedCells.size };
    });
    expect(state.selectedCol).not.toBeNull();
    expect(state.copyWithHeader).toBe(true);
    expect(state.cells).toBeGreaterThan(0);
    // Press Escape — should clear everything via selectNoneCells
    await page.keyboard.press('Escape');
    await page.waitForTimeout(100);
    state = await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'sample1');
      return { selectedCol: w.selectedCol, copyWithHeader: w._copyWithHeader, cells: w.selectedCells.size };
    });
    expect(state.selectedCol).toBeNull();
    expect(state.copyWithHeader).toBe(false);
    expect(state.cells).toBe(0);
  });

  test('Ctrl+A selects all from partial selection', async ({ page }) => {
    // Select a few cells via shift+arrow, then Ctrl+A should select all
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.keyboard.press('Shift+ArrowRight');
    await page.waitForTimeout(100);
    let sel = await getSelection(page, 'sample1');
    expect(sel.cells.length).toBe(2);
    // Ctrl+A should select everything
    await page.keyboard.press('Control+a');
    await page.waitForTimeout(100);
    sel = await getSelection(page, 'sample1');
    expect(sel.cells.length).toBe(30);
  });

  test('copy after row-select then normal select does not include header', async ({ page }) => {
    // Row-select first (sets _copyWithHeader)
    const rowNum = page.locator('.subwindow table tbody td.row-num').first();
    await rowNum.click();
    await page.waitForTimeout(100);
    // Now click a single data cell (should clear _copyWithHeader)
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.waitForTimeout(100);
    await page.keyboard.press('Control+c');
    const clip = await page.evaluate(() => navigator.clipboard.readText());
    expect(clip).toBe('Alice Johnson');
    expect(clip).not.toContain('name');
  });

  // --- Copy from global handler (no cell focus) ---

  test('Ctrl+C works after clicking a column header', async ({ page }) => {
    const header = page.locator('.subwindow table thead th').nth(1);
    await header.click();
    await page.waitForTimeout(100);
    await page.keyboard.press('Control+c');
    const clip = await page.evaluate(() => navigator.clipboard.readText());
    const lines = clip.split('\n');
    expect(lines[0]).toBe('name');
    expect(lines.length).toBe(11); // header + 10 data rows
  });

  test('Ctrl+C works after switching between column headers', async ({ page }) => {
    // Click first column header
    const header1 = page.locator('.subwindow table thead th').nth(1);
    await header1.click();
    await page.waitForTimeout(100);
    // Switch to second column header
    const header2 = page.locator('.subwindow table thead th').nth(2);
    await header2.click();
    await page.waitForTimeout(100);
    await page.keyboard.press('Control+c');
    const clip = await page.evaluate(() => navigator.clipboard.readText());
    const lines = clip.split('\n');
    expect(lines[0]).toBe('email');
    expect(lines[1]).toBe('alice@example.com');
  });

  test('Ctrl+X works after clicking a column header', async ({ page }) => {
    const header = page.locator('.subwindow table thead th').nth(1);
    await header.click();
    await page.waitForTimeout(100);
    await page.keyboard.press('Control+x');
    await page.waitForTimeout(200);
    const clip = await page.evaluate(() => navigator.clipboard.readText());
    const lines = clip.split('\n');
    expect(lines[0]).toBe('name');
    expect(lines.length).toBe(11);
    // Verify cells were cleared
    const data = await getTableData(page, 'sample1');
    expect(data.rows[0].name).toBe('');
    expect(data.rows[1].name).toBe('');
  });

  test('Ctrl+C from titlebar copies previously selected cells', async ({ page }) => {
    // Select all via corner cell
    const corner = page.locator('.subwindow table thead th.row-num-header');
    await corner.click();
    await page.waitForTimeout(100);
    // Click titlebar to move focus away from the table
    await page.locator('.subwindow .win-title').first().click();
    await page.waitForTimeout(100);
    await page.keyboard.press('Control+c');
    const clip = await page.evaluate(() => navigator.clipboard.readText());
    const lines = clip.split('\n');
    expect(lines[0]).toBe('name\temail\tmember_since');
    expect(lines.length).toBe(11);
  });

  test('Ctrl+C global handler does not fire when focus is in an input', async ({ page }) => {
    // Select a cell, then copy to populate clipboard
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.waitForTimeout(100);
    await page.keyboard.press('Control+c');
    const clip1 = await page.evaluate(() => navigator.clipboard.readText());
    expect(clip1).toBe('Alice Johnson');
    // Focus the SQL input — the global handler should NOT fire copySelectedCells
    await page.locator('#sql-input').click();
    await page.keyboard.type('test');
    // Clear clipboard and press Ctrl+C while in the input
    await page.evaluate(() => navigator.clipboard.writeText(''));
    await page.keyboard.press('Control+c');
    const clip2 = await page.evaluate(() => navigator.clipboard.readText());
    // Clipboard should NOT contain the previously selected cell data
    expect(clip2).not.toBe('Alice Johnson');
  });

  // --- Copy formatted vs raw values ---

  test('copy uses formatted display value when transform is active', async ({ page }) => {
    // Load a plugin that transforms the name column
    await page.evaluate(() => {
      const cfg = {
        name: 'Upper Name',
        table: 'sample1',
        columns: [{ match: '^name$', display: 'upper(value)' }]
      };
      const errors = app._test.validatePlugin(cfg);
      if (errors.length) throw new Error(errors.join(', '));
      cfg._compiled = app._test.compilePlugin(cfg);
      app._test.plugins.push(cfg);
      app._test.rebuildTransformCache();
      app._test.rerenderAllWindows();
    });
    await page.waitForTimeout(300);
    // Select the first name cell and copy
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.waitForTimeout(100);
    await page.keyboard.press('Control+c');
    const clip = await page.evaluate(() => navigator.clipboard.readText());
    expect(clip).toBe('ALICE JOHNSON');
  });

  test('copy uses raw value when transform is disabled', async ({ page }) => {
    // Load a plugin that transforms the name column
    await page.evaluate(() => {
      const cfg = {
        name: 'Upper Name',
        table: 'sample1',
        columns: [{ match: '^name$', display: 'upper(value)' }]
      };
      const errors = app._test.validatePlugin(cfg);
      if (errors.length) throw new Error(errors.join(', '));
      cfg._compiled = app._test.compilePlugin(cfg);
      app._test.plugins.push(cfg);
      app._test.rebuildTransformCache();
      app._test.rerenderAllWindows();
    });
    await page.waitForTimeout(300);
    // Disable the transform via Ctrl+Shift+4
    await page.keyboard.press('Control+Shift+4');
    await page.waitForTimeout(200);
    // Select the first name cell and copy
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.waitForTimeout(100);
    await page.keyboard.press('Control+c');
    const clip = await page.evaluate(() => navigator.clipboard.readText());
    expect(clip).toBe('Alice Johnson');
  });

  test('copy multi-cell selection uses formatted values per column', async ({ page }) => {
    // Load a plugin that only transforms the name column
    await page.evaluate(() => {
      const cfg = {
        name: 'Upper Name',
        table: 'sample1',
        columns: [{ match: '^name$', display: 'upper(value)' }]
      };
      const errors = app._test.validatePlugin(cfg);
      if (errors.length) throw new Error(errors.join(', '));
      cfg._compiled = app._test.compilePlugin(cfg);
      app._test.plugins.push(cfg);
      app._test.rebuildTransformCache();
      app._test.rerenderAllWindows();
    });
    await page.waitForTimeout(300);
    // Select name and email of first row
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.keyboard.press('Shift+ArrowRight');
    await page.waitForTimeout(100);
    await page.keyboard.press('Control+c');
    const clip = await page.evaluate(() => navigator.clipboard.readText());
    // name should be formatted (uppercase), email should be raw (no transform)
    expect(clip).toBe('ALICE JOHNSON\talice@example.com');
  });

  test('column header copy uses formatted values when transform is active', async ({ page }) => {
    // Load a plugin that transforms the name column
    await page.evaluate(() => {
      const cfg = {
        name: 'Upper Name',
        table: 'sample1',
        columns: [{ match: '^name$', display: 'upper(value)' }]
      };
      const errors = app._test.validatePlugin(cfg);
      if (errors.length) throw new Error(errors.join(', '));
      cfg._compiled = app._test.compilePlugin(cfg);
      app._test.plugins.push(cfg);
      app._test.rebuildTransformCache();
      app._test.rerenderAllWindows();
    });
    await page.waitForTimeout(300);
    // Click the name column header to select the whole column
    const header = page.locator('.subwindow table thead th').nth(1);
    await header.click();
    await page.waitForTimeout(100);
    await page.keyboard.press('Control+c');
    const clip = await page.evaluate(() => navigator.clipboard.readText());
    const lines = clip.split('\n');
    expect(lines[0]).toBe('name');
    expect(lines[1]).toBe('ALICE JOHNSON');
    expect(lines[2]).toBe('BOB SMITH');
  });

  test('copy shows a toast with cell counts', async ({ page }) => {
    const firstCell = page.locator('.subwindow table tbody td.data-cell').first();
    await firstCell.click();
    await page.keyboard.press('Control+c');
    await expect(page.locator('.toast', { hasText: 'Copied 1 cell' })).toBeVisible();
    await page.evaluate(() => document.querySelectorAll('.toast').forEach(t => t.remove()));
    await page.keyboard.press('Shift+ArrowRight');
    await page.keyboard.press('Shift+ArrowDown');
    await page.keyboard.press('Control+c');
    await expect(page.locator('.toast', { hasText: 'Copied 4 cells (2 rows × 2 columns)' })).toBeVisible();
  });

  test('paste, cut, undo, and redo show descriptive toasts; Ctrl+Y redoes', async ({ page }) => {
    await page.evaluate(() => navigator.clipboard.writeText('X1\tX2'));
    const firstCell = page.locator('.subwindow table tbody td.data-cell').first();
    await firstCell.click();
    await page.keyboard.press('Control+v');
    await expect(page.locator('.toast', { hasText: /Pasted 2 cells \(1 row × 2 columns\) — .+ to undo/ })).toBeVisible();
    await page.evaluate(() => document.querySelectorAll('.toast').forEach(t => t.remove()));

    await page.keyboard.press('Control+z');
    await expect(page.locator('.toast', { hasText: /Undid Paste \(2 cells\) — .+ to redo/ })).toBeVisible();
    await page.evaluate(() => document.querySelectorAll('.toast').forEach(t => t.remove()));

    await page.keyboard.press('Control+y');
    await expect(page.locator('.toast', { hasText: /Redid Paste \(2 cells\) — .+ to undo/ })).toBeVisible();
    let data = await getTableData(page, 'sample1');
    expect(data.rows[0].name).toBe('X1');
    await page.evaluate(() => document.querySelectorAll('.toast').forEach(t => t.remove()));

    await page.keyboard.press('Control+x');
    await expect(page.locator('.toast', { hasText: /Cut 2 cells — .+ to undo/ })).toBeVisible();
    data = await getTableData(page, 'sample1');
    expect(data.rows[0].name).toBe('');
  });
});
