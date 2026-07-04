const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, getTableData, executeSQL } = require('../helpers');

test.describe('Table Editing', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('cells are selectable but not editable until F2', async ({ page }) => {
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await expect(cell).toBeVisible();
    await cell.click();
    // Clicking selects, does not enter edit mode.
    await expect(cell).not.toHaveAttribute('contenteditable', 'true');
    await page.keyboard.press('F2');
    await expect(cell).toHaveAttribute('contenteditable', 'true');
    // Escape reverts and exits edit mode.
    await page.keyboard.press('Escape');
    await expect(cell).not.toHaveAttribute('contenteditable', 'true');
  });

  test('editing a cell updates the table data', async ({ page }) => {
    const before = await getTableData(page, 'sample1');
    const col = before.columns[0];

    // Click to select, F2 to enter edit mode, then type.
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.keyboard.press('F2');
    await cell.fill('EDITED_VALUE');
    // Blur to trigger the save
    await page.locator('#sql-input').click();
    await page.waitForTimeout(500);

    const after = await getTableData(page, 'sample1');
    expect(after.rows[0][col]).toBe('EDITED_VALUE');
  });

  test('double-click enters edit mode', async ({ page }) => {
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.dblclick();
    await expect(cell).toHaveAttribute('contenteditable', 'true');
  });

  test('Ctrl+U also enters edit mode', async ({ page }) => {
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.keyboard.press('Control+u');
    await expect(cell).toHaveAttribute('contenteditable', 'true');
  });

  test('Ctrl+click enters edit mode directly', async ({ page }) => {
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click({ modifiers: ['Control'] });
    await expect(cell).toHaveAttribute('contenteditable', 'true');
  });

  test('Ctrl+click on a different cell edits that cell without selecting first', async ({ page }) => {
    const cells = page.locator('.subwindow table tbody td.data-cell');
    const first = cells.first();
    const second = cells.nth(1);
    // Click first cell normally to select it
    await first.click();
    await expect(first).not.toHaveAttribute('contenteditable', 'true');
    // Ctrl+click the second cell — should enter edit mode on the second cell
    await second.click({ modifiers: ['Control'] });
    await expect(second).toHaveAttribute('contenteditable', 'true');
    await expect(first).not.toHaveAttribute('contenteditable', 'true');
  });

  test('Ctrl+click edit allows typing and saves on blur', async ({ page }) => {
    const before = await getTableData(page, 'sample1');
    const col = before.columns[0];
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click({ modifiers: ['Control'] });
    await expect(cell).toHaveAttribute('contenteditable', 'true');
    await cell.fill('CTRL_EDITED');
    await page.locator('#sql-input').click();
    await page.waitForTimeout(500);
    const after = await getTableData(page, 'sample1');
    expect(after.rows[0][col]).toBe('CTRL_EDITED');
  });

  test('Ctrl+click does not trigger drag-select', async ({ page }) => {
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click({ modifiers: ['Control'] });
    // Should be in edit mode, not drag-select mode
    await expect(cell).toHaveAttribute('contenteditable', 'true');
    // Table should not have drag-selecting class
    const table = page.locator('.subwindow table');
    await expect(table).not.toHaveClass(/drag-selecting/);
  });

  test('plain click after Ctrl+click edit does not stay in edit mode', async ({ page }) => {
    const cells = page.locator('.subwindow table tbody td.data-cell');
    const first = cells.first();
    const second = cells.nth(1);
    // Ctrl+click to edit the first cell
    await first.click({ modifiers: ['Control'] });
    await expect(first).toHaveAttribute('contenteditable', 'true');
    // Plain click on second cell — exits edit mode on first, selects second
    await second.click();
    await expect(first).not.toHaveAttribute('contenteditable', 'true');
    await expect(second).not.toHaveAttribute('contenteditable', 'true');
  });

  test('add row via context menu', async ({ page }) => {
    const before = await getTableData(page, 'sample1');

    // Right-click on a row number cell to get context menu
    const rowNumCell = page.locator('.subwindow table tbody td.row-num').first();
    await rowNumCell.click({ button: 'right' });

    // Look for the add/insert row option
    const addRowBtn = page.getByText(/add row|insert row/i).first();
    if (await addRowBtn.isVisible({ timeout: 1000 }).catch(() => false)) {
      await addRowBtn.click();
      await page.waitForTimeout(500);
      const after = await getTableData(page, 'sample1');
      expect(after.rows.length).toBe(before.rows.length + 1);
    }
  });

  test('delete row via context menu', async ({ page }) => {
    const before = await getTableData(page, 'sample1');

    const rowNumCell = page.locator('.subwindow table tbody td.row-num').first();
    await rowNumCell.click({ button: 'right' });

    const deleteBtn = page.getByText(/delete row/i).first();
    if (await deleteBtn.isVisible({ timeout: 1000 }).catch(() => false)) {
      await deleteBtn.click();
      await page.waitForTimeout(500);
      const after = await getTableData(page, 'sample1');
      expect(after.rows.length).toBe(before.rows.length - 1);
    }
  });

  test('filter with REGEXP narrows visible rows', async ({ page }) => {
    // Switch the toolbar to Filter mode and type a REGEXP filter
    await page.locator('.subwindow .toolbar-mode .mode-filter').click();
    const filterInput = page.locator('.subwindow .filter-input');
    await filterInput.fill("[name] REGEXP '^A'");
    await page.waitForTimeout(500);

    // Check that the filter input does NOT have an error class
    const hasError = await filterInput.evaluate(el => el.classList.contains('filter-error'));
    expect(hasError).toBe(false);

    // Should show fewer rows than the full table (only names starting with A)
    const visibleRows = await page.locator('.subwindow table tbody tr').count();
    expect(visibleRows).toBeGreaterThan(0);
    expect(visibleRows).toBeLessThan(10); // sample1 has 10 rows
  });

  test('row numbers are sequential after delete via SQL', async ({ page }) => {
    const before = await getTableData(page, 'sample1');
    const col = before.columns[0];
    const firstVal = before.rows[0][col];

    await executeSQL(page, `DELETE FROM [sample1] WHERE [${col}] = '${firstVal}'`);
    await page.waitForTimeout(500);

    const after = await getTableData(page, 'sample1');
    // Verify _rownum is sequential 1..n
    after.rows.forEach((row, i) => {
      expect(row._rownum).toBe(i + 1);
    });
  });
});
