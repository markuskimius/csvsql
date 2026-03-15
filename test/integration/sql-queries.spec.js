const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, executeSQL, getTableData } = require('../helpers');

test.describe('SQL Queries', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('SELECT * returns all rows', async ({ page }) => {
    const status = await executeSQL(page, 'SELECT * FROM [sample1]');
    expect(status).toContain('row');

    // A query result window should appear
    const windowCount = await page.evaluate(() => app._test.windows.length);
    expect(windowCount).toBe(2); // original + query result
  });

  test('SELECT with WHERE filters rows', async ({ page }) => {
    const originalData = await getTableData(page, 'sample1');
    const col = originalData.columns[0];
    const firstVal = originalData.rows[0][col];

    const status = await executeSQL(page, `SELECT * FROM [sample1] WHERE [${col}] = '${firstVal}'`);
    expect(status).toContain('row');
  });

  test('INSERT adds a row', async ({ page }) => {
    const before = await getTableData(page, 'sample1');
    const cols = before.columns;
    const values = cols.map(() => "'test'").join(', ');

    const status = await executeSQL(page, `INSERT INTO [sample1] (${cols.map(c => `[${c}]`).join(', ')}) VALUES (${values})`);
    expect(status).toMatch(/affected|OK/i);

    // DML syncs back via refreshAllTableWindows
    const after = await getTableData(page, 'sample1');
    expect(after.rows.length).toBe(before.rows.length + 1);
  });

  test('DELETE removes rows', async ({ page }) => {
    const before = await getTableData(page, 'sample1');
    const col = before.columns[0];
    const firstVal = before.rows[0][col];

    const status = await executeSQL(page, `DELETE FROM [sample1] WHERE [${col}] = '${firstVal}'`);
    expect(status).toMatch(/affected|OK/i);

    const after = await getTableData(page, 'sample1');
    expect(after.rows.length).toBeLessThan(before.rows.length);
  });

  test('UPDATE modifies cell values', async ({ page }) => {
    const before = await getTableData(page, 'sample1');
    const col = before.columns[0];

    const status = await executeSQL(page, `UPDATE [sample1] SET [${col}] = 'UPDATED' WHERE rowid = 1`);
    expect(status).toMatch(/affected|OK/i);

    const after = await getTableData(page, 'sample1');
    expect(after.rows[0][col]).toBe('UPDATED');
  });

  test('CREATE TABLE creates a new table', async ({ page }) => {
    const status = await executeSQL(page, 'CREATE TABLE [newtbl] (id TEXT, name TEXT)');
    expect(status).toContain('newtbl');

    // Wait for the new window to appear
    await waitForWindow(page, 'newtbl');

    const data = await getTableData(page, 'newtbl');
    expect(data).not.toBeNull();
    // Empty table may have no columns detected via SELECT *
  });

  test('SQL syntax error shows error message', async ({ page }) => {
    const status = await executeSQL(page, 'SELECTT * FROMM nowhere');
    expect(status.toLowerCase()).toMatch(/error/);
  });
});
