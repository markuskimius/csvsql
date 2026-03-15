const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, getTableData } = require('../helpers');

test.describe('Save', () => {
  test('save produces valid CSV via Papa.unparse', async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');

    // Get the CSV output that would be saved
    const csv = await page.evaluate(() => {
      const t = app._test.tables['sample1'];
      // Replicate what save does: unparse with columns + rows (minus _rownum)
      const rows = t.rows.map(r => {
        const obj = {};
        t.columns.forEach(c => { obj[c] = r[c]; });
        return obj;
      });
      return Papa.unparse({ fields: t.columns, data: rows });
    });

    expect(csv).toBeTruthy();
    // Should have header line + data lines
    const lines = csv.trim().split('\n');
    const data = await getTableData(page, 'sample1');
    expect(lines.length).toBe(data.rows.length + 1); // header + rows
  });

  test('round-trip: parse → unparse preserves data', async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');

    const roundTrip = await page.evaluate(() => {
      const t = app._test.tables['sample1'];
      const rows = t.rows.map(r => {
        const obj = {};
        t.columns.forEach(c => { obj[c] = r[c]; });
        return obj;
      });
      const csv = Papa.unparse({ fields: t.columns, data: rows });

      // Re-parse
      const parsed = Papa.parse(csv, { header: true, skipEmptyLines: true });
      return {
        originalCols: t.columns,
        parsedCols: parsed.meta.fields,
        originalRowCount: t.rows.length,
        parsedRowCount: parsed.data.length,
        firstOriginal: rows[0],
        firstParsed: parsed.data[0],
      };
    });

    expect(roundTrip.parsedCols).toEqual(roundTrip.originalCols);
    expect(roundTrip.parsedRowCount).toBe(roundTrip.originalRowCount);
    expect(roundTrip.firstParsed).toEqual(roundTrip.firstOriginal);
  });
});
