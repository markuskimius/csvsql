const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow } = require('../helpers');

// All ids parse to the same double (1e19 region) except 2, so correct
// ordering requires the exact BigInt tie-break in the sort comparator.
test.describe('Numeric sort precision', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/bigint.csv');
    await waitForWindow(page, 'bigint');
  });

  async function sortValues(page, col, dir) {
    return page.evaluate(({ col, dir }) => {
      const win = app._test.windows[0];
      win.sortCols = [{ col, dir }];
      app._test.rebuildTable(win);
      return win._displayRows.map(r => r[col]);
    }, { col, dir });
  }
  const sortIds = (page, dir) => sortValues(page, 'id', dir);

  test('integers beyond 2^53 sort by exact value ascending', async ({ page }) => {
    expect(await sortIds(page, 'asc')).toEqual([
      '-10000000000000000001',
      '2',
      '9999999999999999999',
      '10000000000000000000',
      '10000000000000000001',
    ]);
  });

  test('integers beyond 2^53 sort by exact value descending', async ({ page }) => {
    expect(await sortIds(page, 'desc')).toEqual([
      '10000000000000000001',
      '10000000000000000000',
      '9999999999999999999',
      '2',
      '-10000000000000000001',
    ]);
  });

  test('mixed numbers, text, and empty cells sort deterministically', async ({ page }) => {
    // Numbers order numerically among themselves; empties first, text after
    // (collator fallback for non-numeric pairs)
    expect(await sortValues(page, 'mixed', 'asc')).toEqual(['', '2', '9', '10', 'abc']);
    expect(await sortValues(page, 'mixed', 'desc')).toEqual(['abc', '10', '9', '2', '']);
  });

  test('plain text column sorts alphabetically', async ({ page }) => {
    expect(await sortValues(page, 'label', 'asc')).toEqual(['a', 'b', 'c', 'd', 'e']);
    expect(await sortValues(page, 'label', 'desc')).toEqual(['e', 'd', 'c', 'b', 'a']);
  });

  test('autofilter value list orders numerically', async ({ page }) => {
    const th = page.locator('.data-table th:not(.row-num-header)').first();
    await th.locator('.col-filter-btn').click();
    const dropdown = page.locator('.autofilter-dropdown');
    await expect(dropdown).toBeVisible();

    const labels = await dropdown.locator('.autofilter-item').allTextContents();
    expect(labels.map(s => s.trim())).toEqual([
      '-10000000000000000001',
      '2',
      '9999999999999999999',
      '10000000000000000000',
      '10000000000000000001',
    ]);
  });
});

test.describe('Autofilter ordering with display transforms', () => {
  test('currency-formatted column still orders by numeric raw value', async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../example/orders.csv');
    await waitForWindow(page, 'orders');

    await page.evaluate(() => {
      const cfg = {
        name: 'USD',
        table: '.*',
        columns: [{ match: '^total$', display: "isNum(value) ? '$' + fixed(num(value), 2) : value" }]
      };
      const errors = app._test.validatePlugin(cfg);
      if (errors.length) throw new Error(errors.join(', '));
      cfg._compiled = app._test.compilePlugin(cfg);
      app._test.plugins.push(cfg);
      app._test.rebuildTransformCache();
    });

    const th = page.locator('.data-table th:not(.row-num-header)', { hasText: 'total' });
    await th.locator('.col-filter-btn').click();
    const dropdown = page.locator('.autofilter-dropdown');
    await expect(dropdown).toBeVisible();

    // Displayed formatted, but ordered by the raw numeric value —
    // lexical order would wrongly put $129.50 first
    const labels = await dropdown.locator('.autofilter-item').allTextContents();
    expect(labels.map(s => s.trim())).toEqual(['$49.99', '$89.00', '$129.50']);
  });
});
