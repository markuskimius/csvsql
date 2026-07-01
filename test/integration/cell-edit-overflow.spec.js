const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, executeSQL } = require('../helpers');

// ---------------------------------------------------------------------------
// Auto-fit measures ALL rows (not just viewport)
// ---------------------------------------------------------------------------

test.describe('Auto-fit measures all rows', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('auto-fit on a small table sizes to the widest cell', async ({ page }) => {
    // Set a known long value in one cell so we can verify auto-fit finds it
    await page.evaluate(() => {
      const t = app._test.tables.sample1;
      t.rows[0][t.columns[0]] = 'A_VERY_LONG_VALUE_THAT_IS_WIDER_THAN_ANYTHING_ELSE';
      app._test.rebuildTable(app._test.windows[0]);
    });
    await page.waitForTimeout(300);

    // Record the current width, then set it artificially narrow
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.colWidths[0] = 50;
      win._colgroup.children[1].style.width = '50px';
    });

    // Auto-fit all columns
    await page.evaluate(() => {
      const win = app._test.windows[0];
      app._test.autoFitAllColumns(win);
    });
    await page.waitForTimeout(300);

    const width = await page.evaluate(() => app._test.windows[0].colWidths[0]);
    // The long value should cause auto-fit to produce a width well above 50px
    expect(width).toBeGreaterThan(100);
  });

  test('auto-fit on a large table finds widest value outside the viewport', async ({ page }) => {
    // Create a 200-row table by duplicating sample1 rows, then inject a long value at row 150
    await executeSQL(page, "SELECT [name] INTO [big_table] FROM [sample1]");
    await waitForWindow(page, 'big_table');
    await page.waitForTimeout(300);

    // Expand the table to 200 rows and place a long value at row 150
    const longVal = 'THIS_IS_AN_EXTREMELY_LONG_VALUE_THAT_EXCEEDS_ANYTHING_IN_THE_VIEWPORT_AREA_BY_FAR';
    await page.evaluate((longVal) => {
      const t = app._test.tables.big_table;
      // Pad to 200 rows
      while (t.rows.length < 200) {
        t.rows.push({ name: 'short', _rownum: t.rows.length + 1 });
      }
      t.rows[149].name = longVal;
      const win = app._test.windows.find(w => w.tableName === 'big_table');
      app._test.rebuildTable(win);
    }, longVal);
    await page.waitForTimeout(500);

    // Verify that row 150 is NOT in the rendered viewport
    const renderInfo = await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'big_table');
      return { start: win._renderStart, end: win._renderEnd, total: win._displayRows.length };
    });
    expect(renderInfo.total).toBe(200);
    expect(renderInfo.end).toBeLessThan(150);

    // Set column width to something narrow
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'big_table');
      win.colWidths[0] = 50;
      win._colgroup.children[1].style.width = '50px';
    });

    // Auto-fit should measure ALL rows including the off-viewport long one
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'big_table');
      app._test.autoFitAllColumns(win);
    });
    await page.waitForTimeout(300);

    const fittedWidth = await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'big_table');
      return win.colWidths[0];
    });

    // The fitted width must be large enough for the long value (well over 100px)
    expect(fittedWidth).toBeGreaterThan(200);
  });

  test('auto-fit on a large table produces wider column than viewport-only measurement would', async ({ page }) => {
    // Create a table from sample1, then expand to 300 rows with a long value at row 250
    await executeSQL(page, "SELECT [name] INTO [measure_test] FROM [sample1]");
    await waitForWindow(page, 'measure_test');
    await page.waitForTimeout(300);

    const longVal = 'AAAA_BBBB_CCCC_DDDD_EEEE_FFFF_GGGG_HHHH_IIII_JJJJ';
    await page.evaluate((longVal) => {
      const t = app._test.tables.measure_test;
      while (t.rows.length < 300) {
        t.rows.push({ name: 'x', _rownum: t.rows.length + 1 });
      }
      t.rows[249].name = longVal;
      const win = app._test.windows.find(w => w.tableName === 'measure_test');
      app._test.rebuildTable(win);
    }, longVal);
    await page.waitForTimeout(500);

    // Set narrow width
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'measure_test');
      win.colWidths[0] = 40;
      win._colgroup.children[1].style.width = '40px';
    });

    // Auto-fit
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'measure_test');
      app._test.autoFitAllColumns(win);
    });
    await page.waitForTimeout(300);

    const fittedWidth = await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'measure_test');
      return win.colWidths[0];
    });

    // If only viewport rows were measured, width would be small (just for 'x').
    // With all-row measurement, it must accommodate the long value.
    expect(fittedWidth).toBeGreaterThan(100);
  });
});

// ---------------------------------------------------------------------------
// Cell edit mode overflow CSS
// ---------------------------------------------------------------------------

test.describe('Cell edit mode overflow CSS', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('non-edit cell has overflow hidden', async ({ page }) => {
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    const overflow = await cell.evaluate(el => getComputedStyle(el).overflow);
    expect(overflow).toBe('hidden');
  });

  test('non-edit cell has text-overflow ellipsis', async ({ page }) => {
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    const textOverflow = await cell.evaluate(el => getComputedStyle(el).textOverflow);
    expect(textOverflow).toBe('ellipsis');
  });

  test('edit mode cell has overflow auto', async ({ page }) => {
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.keyboard.press('F2');
    await expect(cell).toHaveAttribute('contenteditable', 'true');

    const overflow = await cell.evaluate(el => getComputedStyle(el).overflow);
    expect(overflow).toBe('auto');
  });

  test('edit mode cell has text-overflow clip', async ({ page }) => {
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.keyboard.press('F2');
    await expect(cell).toHaveAttribute('contenteditable', 'true');

    const textOverflow = await cell.evaluate(el => getComputedStyle(el).textOverflow);
    expect(textOverflow).toBe('clip');
  });

  test('edit mode cell has hidden scrollbars via scrollbar-width none', async ({ page }) => {
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.keyboard.press('F2');
    await expect(cell).toHaveAttribute('contenteditable', 'true');

    const scrollbarWidth = await cell.evaluate(el => getComputedStyle(el).scrollbarWidth);
    expect(scrollbarWidth).toBe('none');
  });

  test('exiting edit mode restores overflow hidden', async ({ page }) => {
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.keyboard.press('F2');
    await expect(cell).toHaveAttribute('contenteditable', 'true');

    // Exit edit mode via Escape
    await page.keyboard.press('Escape');
    await expect(cell).not.toHaveAttribute('contenteditable', 'true');

    const overflow = await cell.evaluate(el => getComputedStyle(el).overflow);
    expect(overflow).toBe('hidden');
  });
});

// ---------------------------------------------------------------------------
// Cell scroll position management
// ---------------------------------------------------------------------------

test.describe('Cell scroll position management', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('entering edit mode on a cell with long text scrolls to the end', async ({ page }) => {
    // First, put a very long value in a cell that will overflow
    const longText = 'ABCDEFGHIJ_1234567890_ABCDEFGHIJ_1234567890_ABCDEFGHIJ_1234567890_END';
    await page.evaluate((text) => {
      const t = app._test.tables.sample1;
      t.rows[0][t.columns[0]] = text;
      app._test.rebuildTable(app._test.windows[0]);
    }, longText);
    await page.waitForTimeout(300);

    // Make the column narrow so text overflows
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.colWidths[0] = 80;
      win._colgroup.children[1].style.width = '80px';
      app._test.updateTableWidth(win);
    });
    await page.waitForTimeout(200);

    // Click the cell and enter edit mode
    const cell = page.locator('.subwindow table tbody tr:first-child td.data-cell').first();
    await cell.click();
    await page.keyboard.press('F2');
    await expect(cell).toHaveAttribute('contenteditable', 'true');
    await page.waitForTimeout(200);

    // scrollLeft should be > 0 because enterEditMode scrolls to the end
    const scrollLeft = await cell.evaluate(el => el.scrollLeft);
    expect(scrollLeft).toBeGreaterThan(0);
  });

  test('exiting edit mode resets scrollLeft to 0', async ({ page }) => {
    // Put a long value in a cell
    const longText = 'XXXXXXXXXX_YYYYYYYYYY_ZZZZZZZZZZ_1234567890_ABCDEFGHIJ_KLMNOPQRST';
    await page.evaluate((text) => {
      const t = app._test.tables.sample1;
      t.rows[0][t.columns[0]] = text;
      app._test.rebuildTable(app._test.windows[0]);
    }, longText);
    await page.waitForTimeout(300);

    // Narrow the column
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.colWidths[0] = 80;
      win._colgroup.children[1].style.width = '80px';
      app._test.updateTableWidth(win);
    });
    await page.waitForTimeout(200);

    // Enter edit mode
    const cell = page.locator('.subwindow table tbody tr:first-child td.data-cell').first();
    await cell.click();
    await page.keyboard.press('F2');
    await expect(cell).toHaveAttribute('contenteditable', 'true');
    await page.waitForTimeout(200);

    // Verify scrollLeft > 0 in edit mode
    const scrollLeftInEdit = await cell.evaluate(el => el.scrollLeft);
    expect(scrollLeftInEdit).toBeGreaterThan(0);

    // Exit edit mode via Escape
    await page.keyboard.press('Escape');
    await expect(cell).not.toHaveAttribute('contenteditable', 'true');
    await page.waitForTimeout(200);

    // scrollLeft should be reset to 0
    const scrollLeftAfterExit = await cell.evaluate(el => el.scrollLeft);
    expect(scrollLeftAfterExit).toBe(0);
  });

  test('exiting edit mode via blur resets scrollLeft to 0', async ({ page }) => {
    // Put a long value in a cell
    const longText = 'AAAAAAAAA_BBBBBBBBB_CCCCCCCCC_DDDDDDDDD_EEEEEEEEE_FFFFFFFFF_GGGGG';
    await page.evaluate((text) => {
      const t = app._test.tables.sample1;
      t.rows[0][t.columns[0]] = text;
      app._test.rebuildTable(app._test.windows[0]);
    }, longText);
    await page.waitForTimeout(300);

    // Narrow the column
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.colWidths[0] = 80;
      win._colgroup.children[1].style.width = '80px';
      app._test.updateTableWidth(win);
    });
    await page.waitForTimeout(200);

    // Enter edit mode
    const cell = page.locator('.subwindow table tbody tr:first-child td.data-cell').first();
    await cell.click();
    await page.keyboard.press('F2');
    await expect(cell).toHaveAttribute('contenteditable', 'true');
    await page.waitForTimeout(200);

    // Blur by clicking the SQL input
    await page.locator('#sql-input').click();
    await page.waitForTimeout(300);

    // scrollLeft should be 0 after blur
    const scrollLeft = await cell.evaluate(el => el.scrollLeft);
    expect(scrollLeft).toBe(0);
  });

  test('typing text that overflows the cell keeps caret visible via scrollLeft', async ({ page }) => {
    // Start with an empty cell
    await page.evaluate(() => {
      const t = app._test.tables.sample1;
      t.rows[0][t.columns[0]] = '';
      app._test.rebuildTable(app._test.windows[0]);
    });
    await page.waitForTimeout(300);

    // Narrow the column
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.colWidths[0] = 80;
      win._colgroup.children[1].style.width = '80px';
      app._test.updateTableWidth(win);
    });
    await page.waitForTimeout(200);

    // Enter edit mode and type a lot of text
    const cell = page.locator('.subwindow table tbody tr:first-child td.data-cell').first();
    await cell.click();
    await page.keyboard.press('F2');
    await expect(cell).toHaveAttribute('contenteditable', 'true');

    // Type enough text to overflow the narrow cell
    await page.keyboard.type('ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890ABCDEFGHIJKLMNOP', { delay: 10 });
    await page.waitForTimeout(200);

    // The cell should have scrolled to keep the caret visible
    const scrollLeft = await cell.evaluate(el => el.scrollLeft);
    expect(scrollLeft).toBeGreaterThan(0);

    // The scrollWidth should exceed the visible width
    const dims = await cell.evaluate(el => ({
      scrollWidth: el.scrollWidth,
      clientWidth: el.clientWidth
    }));
    expect(dims.scrollWidth).toBeGreaterThan(dims.clientWidth);
  });
});
