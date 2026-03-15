const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, getTableData } = require('../helpers');

test.describe('File Open', () => {
  test('opens a CSV file and shows correct data', async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');

    const data = await getTableData(page, 'sample1');
    expect(data).not.toBeNull();
    expect(data.columns.length).toBeGreaterThan(0);
    expect(data.rows.length).toBeGreaterThan(0);
    expect(data.filename).toBe('sample1.csv');
  });

  test('opens a TSV file (gzipped) with correct delimiter', async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample2.tsv.gz');
    await waitForWindow(page, 'sample2');

    const data = await getTableData(page, 'sample2');
    expect(data).not.toBeNull();
    expect(data.columns.length).toBeGreaterThan(0);
    expect(data.rows.length).toBeGreaterThan(0);
  });

  test('opens a PSV file', async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample2.psv');
    await waitForWindow(page, 'sample2');

    const data = await getTableData(page, 'sample2');
    expect(data).not.toBeNull();
    expect(data.rows.length).toBeGreaterThan(0);
  });

  test('opens an Excel file with multiple sheets', async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample.xlsx');
    // Wait a bit for multi-sheet processing
    await page.waitForTimeout(2000);

    const windowCount = await page.evaluate(() => app._test.windows.length);
    expect(windowCount).toBeGreaterThanOrEqual(1);
  });

  test('opens a ZIP file with multiple CSVs', async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample.zip');
    // Wait for extraction and parsing
    await page.waitForTimeout(3000);

    const windowCount = await page.evaluate(() => app._test.windows.length);
    expect(windowCount).toBeGreaterThanOrEqual(2);
  });

  test('opens a gzipped CSV file', async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv.gz');
    await waitForWindow(page, 'sample1');

    const data = await getTableData(page, 'sample1');
    expect(data).not.toBeNull();
    expect(data.rows.length).toBeGreaterThan(0);
  });
});
