const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow } = require('../helpers');

function getColWidths(page) {
  return page.evaluate(() => {
    const win = app._test.windows[0];
    return win.colWidths ? [...win.colWidths] : null;
  });
}

function getRowNumWidth(page) {
  return page.evaluate(() => {
    const win = app._test.windows[0];
    return win.rowNumWidth || 50;
  });
}

test.describe('Auto Fit All Columns', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('auto fit all columns does not make column too narrow for content', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      app._test.autoFitAllColumns(win);
    });
    await page.waitForTimeout(300);

    const truncatedCells = await page.evaluate(() => {
      const truncated = [];
      const cells = document.querySelectorAll('.data-table tbody td.data-cell');
      for (const cell of cells) {
        if (cell.scrollWidth > cell.clientWidth) {
          truncated.push(cell.textContent);
        }
      }
      return truncated;
    });
    expect(truncatedCells).toEqual([]);
  });

  test('auto fit all columns called via _test works correctly', async ({ page }) => {
    const before = await getColWidths(page);

    // Make all columns wide first
    await page.evaluate(() => {
      const win = app._test.windows[0];
      for (let i = 0; i < win.colWidths.length; i++) {
        win.colWidths[i] = 400;
        win._colgroup.children[i + 1].style.width = '400px';
      }
    });

    const widened = await getColWidths(page);
    widened.forEach(w => expect(w).toBe(400));

    await page.evaluate(() => {
      const win = app._test.windows[0];
      app._test.autoFitAllColumns(win);
    });

    const after = await getColWidths(page);
    after.forEach(w => {
      expect(w).toBeLessThan(400);
      expect(w).toBeGreaterThanOrEqual(40);
    });
  });
});

test.describe('Auto Fit All Columns — with display transforms', () => {
  async function loadPluginConfig(page, config) {
    await page.evaluate((cfg) => {
      const errors = app._test.validatePlugin(cfg);
      if (errors.length) throw new Error(errors.join(', '));
      const compiled = app._test.compilePlugin(cfg);
      cfg._compiled = compiled;
      app._test.plugins.push(cfg);
      app._test.rebuildTransformCache();
      app._test.rerenderAllWindows();
    }, config);
  }

  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../example/products.csv');
    await waitForWindow(page, 'products');
  });

  test('auto fit measures formatted values when transform is active', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Wide Price',
      table: 'products',
      columns: [
        { match: '^price$', display: '"Total price is exactly: $" + value + " in United States currency"' }
      ]
    });
    await page.waitForTimeout(300);

    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'products');
      app._test.autoFitAllColumns(win);
    });
    const widthsWithTransform = await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'products');
      return [...win.colWidths];
    });

    const priceIdx = await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'products');
      return win._columns.indexOf('price');
    });
    expect(widthsWithTransform[priceIdx]).toBeGreaterThan(100);
  });

  test('auto fit measures raw values when transform is disabled', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Wide Price',
      table: 'products',
      columns: [
        { match: '^price$', display: '"Total price is exactly: $" + value + " in United States currency"' }
      ]
    });
    await page.waitForTimeout(300);

    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'products');
      app._test.autoFitAllColumns(win);
    });
    const widthsWithTransform = await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'products');
      return [...win.colWidths];
    });

    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'products');
      win.disabledTransforms.add('price');
      app._test.autoFitAllColumns(win);
    });
    const widthsWithoutTransform = await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'products');
      return [...win.colWidths];
    });

    const priceIdx = await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'products');
      return win._columns.indexOf('price');
    });

    expect(widthsWithoutTransform[priceIdx]).toBeLessThan(widthsWithTransform[priceIdx]);
  });
});

test.describe('Auto Fit All Columns — docked windows', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
    await uploadFile(page, '../example/products.csv');
    await waitForWindow(page, 'products');
  });

  test('auto fit uses dock leaf pane width for 75% cap when docked', async ({ page }) => {
    await page.evaluate(() => {
      const w0 = app._test.windows[0];
      const w1 = app._test.windows[1];
      app._test.mergeWindowsAsSplit(w0.id, w1.id, 'right');
    });
    await page.waitForTimeout(500);

    const { colWidths, leafWidth } = await page.evaluate(() => {
      const win = app._test.windows[0];
      app._test.autoFitAllColumns(win);
      const lw = win.dockLeaf.el.getBoundingClientRect().width;
      return { colWidths: [...win.colWidths], leafWidth: lw };
    });

    const maxAllowed = Math.floor(leafWidth * 0.75);
    colWidths.forEach(w => expect(w).toBeLessThanOrEqual(maxAllowed));
  });
});

test.describe('Column header context menu', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('right-click on column header shows context menu', async ({ page }) => {
    const th = page.locator('.data-table th:not(.row-num-header)').first();
    await th.click({ button: 'right' });
    await page.waitForTimeout(300);
    await expect(page.locator('.context-menu')).toHaveCount(1);
    await expect(page.locator('.context-menu button', { hasText: 'Insert Column Right' })).toBeVisible();
    await expect(page.locator('.context-menu button', { hasText: 'Delete Column' })).toBeVisible();
  });
});
