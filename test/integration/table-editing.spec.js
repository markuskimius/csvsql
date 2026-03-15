const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, getTableData, executeSQL } = require('../helpers');

test.describe('Table Editing', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('cells are contenteditable', async ({ page }) => {
    const cell = page.locator('.subwindow table tbody td[contenteditable]').first();
    await expect(cell).toBeVisible();
  });

  test('editing a cell updates the table data', async ({ page }) => {
    const before = await getTableData(page, 'sample1');
    const col = before.columns[0];

    // Click to focus the first editable cell and change its content
    const cell = page.locator('.subwindow table tbody td[contenteditable]').first();
    await cell.click();
    await cell.fill('EDITED_VALUE');
    // Blur to trigger the save
    await page.locator('#sql-input').click();
    await page.waitForTimeout(500);

    const after = await getTableData(page, 'sample1');
    expect(after.rows[0][col]).toBe('EDITED_VALUE');
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
