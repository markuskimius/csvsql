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

async function openFilterDropdown(page, colIndex) {
  const th = page.locator('.data-table th:not(.row-num-header)').nth(colIndex);
  await th.locator('.col-filter-btn').click();
  await page.waitForTimeout(300);
  return page.locator('.autofilter-dropdown');
}

test.describe('Auto Fit Columns — filter dropdown buttons', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('autofilter dropdown contains sizing section with two buttons', async ({ page }) => {
    const dropdown = await openFilterDropdown(page, 0);
    const sizing = dropdown.locator('.autofilter-sizing');
    await expect(sizing).toBeVisible();

    const buttons = sizing.locator('button');
    expect(await buttons.count()).toBe(2);
    expect(await buttons.nth(0).textContent()).toBe('Auto Fit This Column');
    expect(await buttons.nth(1).textContent()).toBe('Auto Fit All Columns');
  });

  test('sizing section has a border-top divider', async ({ page }) => {
    const dropdown = await openFilterDropdown(page, 0);
    const sizing = dropdown.locator('.autofilter-sizing');
    const borderTop = await sizing.evaluate(el => getComputedStyle(el).borderTopStyle);
    expect(borderTop).toBe('solid');
  });

  test('sizing section appears below the Apply/Clear buttons', async ({ page }) => {
    const dropdown = await openFilterDropdown(page, 0);
    const buttonsBox = await dropdown.locator('.autofilter-buttons').boundingBox();
    const sizingBox = await dropdown.locator('.autofilter-sizing').boundingBox();
    expect(sizingBox.y).toBeGreaterThan(buttonsBox.y);
  });
});

test.describe('Auto Fit This Column — via filter dropdown', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('auto fit this column resizes only the clicked column', async ({ page }) => {
    const before = await getColWidths(page);

    // Make column 0 very wide
    const handle = page.locator('th:not(.row-num-header) .col-resize-handle').first();
    const box = await handle.boundingBox();
    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 252, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const widened = await getColWidths(page);
    expect(widened[0]).toBeGreaterThan(before[0] + 200);

    // Click Auto Fit This Column from filter dropdown
    const dropdown = await openFilterDropdown(page, 0);
    await dropdown.locator('.autofilter-sizing button', { hasText: 'Auto Fit This Column' }).click();
    await page.waitForTimeout(300);

    const after = await getColWidths(page);
    expect(after[0]).toBeLessThan(widened[0]);
    expect(after[0]).toBeGreaterThanOrEqual(40);
    // Other columns unchanged
    expect(after[1]).toBe(widened[1]);
    expect(after[2]).toBe(widened[2]);
  });

  test('auto fit this column closes the dropdown', async ({ page }) => {
    const dropdown = await openFilterDropdown(page, 0);
    await dropdown.locator('.autofilter-sizing button', { hasText: 'Auto Fit This Column' }).click();
    await page.waitForTimeout(300);
    await expect(page.locator('.autofilter-dropdown')).toHaveCount(0);
  });

  test('auto fit this column respects minimum width', async ({ page }) => {
    const dropdown = await openFilterDropdown(page, 0);
    await dropdown.locator('.autofilter-sizing button', { hasText: 'Auto Fit This Column' }).click();
    await page.waitForTimeout(300);

    const widths = await getColWidths(page);
    widths.forEach(w => expect(w).toBeGreaterThanOrEqual(40));
  });
});

test.describe('Auto Fit All Columns', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('auto fit all columns resizes every column', async ({ page }) => {
    // Make column 0 very wide
    const handle = page.locator('th:not(.row-num-header) .col-resize-handle').first();
    const box = await handle.boundingBox();
    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 252, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const widened = await getColWidths(page);

    // Click Auto Fit All Columns from filter dropdown
    const dropdown = await openFilterDropdown(page, 0);
    await dropdown.locator('.autofilter-sizing button', { hasText: 'Auto Fit All Columns' }).click();
    await page.waitForTimeout(300);

    const after = await getColWidths(page);
    // Column 0 should shrink from the widened state
    expect(after[0]).toBeLessThan(widened[0]);
    // All columns should be at least MIN_COL_WIDTH
    after.forEach(w => expect(w).toBeGreaterThanOrEqual(40));
  });

  test('auto fit all columns also fits the row-number column', async ({ page }) => {
    // Widen the row-num column
    const handle = page.locator('th.row-num-header .col-resize-handle');
    const box = await handle.boundingBox();
    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 152, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const widenedRowNum = await getRowNumWidth(page);
    expect(widenedRowNum).toBeGreaterThan(150);

    // Click Auto Fit All Columns
    const dropdown = await openFilterDropdown(page, 0);
    await dropdown.locator('.autofilter-sizing button', { hasText: 'Auto Fit All Columns' }).click();
    await page.waitForTimeout(300);

    const afterRowNum = await getRowNumWidth(page);
    expect(afterRowNum).toBeLessThan(widenedRowNum);
    expect(afterRowNum).toBeGreaterThanOrEqual(30);
  });

  test('auto fit all columns closes the dropdown', async ({ page }) => {
    const dropdown = await openFilterDropdown(page, 0);
    await dropdown.locator('.autofilter-sizing button', { hasText: 'Auto Fit All Columns' }).click();
    await page.waitForTimeout(300);
    await expect(page.locator('.autofilter-dropdown')).toHaveCount(0);
  });

  test('auto fit all columns respects 75% window width cap', async ({ page }) => {
    const dropdown = await openFilterDropdown(page, 0);
    await dropdown.locator('.autofilter-sizing button', { hasText: 'Auto Fit All Columns' }).click();
    await page.waitForTimeout(300);

    const { colWidths, winWidth } = await page.evaluate(() => {
      const win = app._test.windows[0];
      return {
        colWidths: [...win.colWidths],
        winWidth: win.el.offsetWidth
      };
    });

    const maxAllowed = Math.floor(winWidth * 0.75);
    colWidths.forEach(w => expect(w).toBeLessThanOrEqual(maxAllowed));
  });

  test('auto fit all columns updates table width correctly', async ({ page }) => {
    const dropdown = await openFilterDropdown(page, 0);
    await dropdown.locator('.autofilter-sizing button', { hasText: 'Auto Fit All Columns' }).click();
    await page.waitForTimeout(300);

    const { tableWidth, expectedWidth } = await page.evaluate(() => {
      const win = app._test.windows[0];
      const rowNumW = win.rowNumWidth || 50;
      const colSum = win.colWidths.reduce((a, b) => a + b, 0);
      return {
        tableWidth: parseInt(win._table.style.width),
        expectedWidth: rowNumW + colSum
      };
    });
    expect(tableWidth).toBe(expectedWidth);
  });

  test('auto fit all columns does not make column too narrow for content', async ({ page }) => {
    // Auto fit and check no visible cell is truncated (scrollWidth > clientWidth)
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
    // Test the function directly via app._test
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
    // Formatted text is ~60 chars in monospace — should be well over 100px
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

    // Auto fit with transform active
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'products');
      app._test.autoFitAllColumns(win);
    });
    const widthsWithTransform = await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'products');
      return [...win.colWidths];
    });

    // Disable transform on price column only, then auto fit again
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

    // Price column should be narrower without the long formatted text
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
    // Dock the two windows into a split
    await page.evaluate(() => {
      const w0 = app._test.windows[0];
      const w1 = app._test.windows[1];
      app._test.mergeWindowsAsSplit(w0.id, w1.id, 'right');
    });
    await page.waitForTimeout(500);

    // Auto fit in the docked window
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

test.describe('Auto Fit — right-click still shows browser menu', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('right-click on column header does not show custom context menu', async ({ page }) => {
    const th = page.locator('.data-table th:not(.row-num-header)').first();
    await th.click({ button: 'right' });
    await page.waitForTimeout(300);
    // No custom context-menu element should appear
    await expect(page.locator('.context-menu')).toHaveCount(0);
  });
});
