const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, executeSQL, getTableData } = require('../helpers');

test.describe('SELECT INTO', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('creates a new table from query results', async ({ page }) => {
    await executeSQL(page, 'SELECT * INTO [result_table] FROM [sample1]');
    await waitForWindow(page, 'result_table');

    const original = await getTableData(page, 'sample1');
    const result = await getTableData(page, 'result_table');

    expect(result).not.toBeNull();
    expect(result.columns).toEqual(original.columns);
    expect(result.rows.length).toBe(original.rows.length);
  });

  test('works with filtered SELECT INTO', async ({ page }) => {
    const original = await getTableData(page, 'sample1');
    const col = original.columns[0];
    const firstVal = original.rows[0][col];

    await executeSQL(page, `SELECT * INTO [filtered] FROM [sample1] WHERE [${col}] = '${firstVal}'`);
    await waitForWindow(page, 'filtered');

    const result = await getTableData(page, 'filtered');
    expect(result).not.toBeNull();
    expect(result.rows.length).toBeGreaterThan(0);
    expect(result.rows.length).toBeLessThanOrEqual(original.rows.length);
  });
});
