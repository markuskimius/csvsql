const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, getTableData } = require('../helpers');

function getColWidths(page) {
  return page.evaluate(() => {
    const win = app._test.windows[0];
    return win.colWidths ? [...win.colWidths] : null;
  });
}

test.describe('Column resize', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('table gets colgroup and fixed-layout on load', async ({ page }) => {
    const hasColgroup = await page.evaluate(() => !!document.querySelector('.data-table colgroup'));
    expect(hasColgroup).toBe(true);

    const isFixed = await page.evaluate(() =>
      document.querySelector('.data-table').classList.contains('fixed-layout'));
    expect(isFixed).toBe(true);
  });

  test('colWidths are set on the window object after load', async ({ page }) => {
    const widths = await getColWidths(page);
    expect(widths).not.toBeNull();
    expect(widths.length).toBe(3);
    widths.forEach(w => expect(w).toBeGreaterThan(0));
  });

  test('right-side resize handles exist on each column header and row-num header', async ({ page }) => {
    const handleCount = await page.evaluate(() =>
      document.querySelectorAll('.col-resize-handle:not(.col-resize-left)').length);
    const colCount = await page.evaluate(() => app._test.tables.sample1.columns.length);
    expect(handleCount).toBe(colCount + 1); // +1 for the row-num header
  });

  test('left-side resize handles exist on each data column header', async ({ page }) => {
    const leftHandleCount = await page.evaluate(() =>
      document.querySelectorAll('.col-resize-handle.col-resize-left').length);
    const colCount = await page.evaluate(() => app._test.tables.sample1.columns.length);
    expect(leftHandleCount).toBe(colCount);
  });

  test('resize handle has col-resize cursor', async ({ page }) => {
    const cursor = await page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first().evaluate(
      el => getComputedStyle(el).cursor);
    expect(cursor).toBe('col-resize');
  });

  test('left-side resize handle has col-resize cursor', async ({ page }) => {
    const cursor = await page.locator('th:not(.row-num-header) .col-resize-handle.col-resize-left').first().evaluate(
      el => getComputedStyle(el).cursor);
    expect(cursor).toBe('col-resize');
  });

  test('dragging a right-side resize handle changes that column width', async ({ page }) => {
    const before = await getColWidths(page);
    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 102, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const after = await getColWidths(page);
    expect(after[0]).toBeGreaterThan(before[0] + 50);
  });

  test('dragging a left-side resize handle changes the previous column width', async ({ page }) => {
    const before = await getColWidths(page);
    // Left handle on the second data column (index 1) should resize column 0
    const handle = page.locator('th:not(.row-num-header) .col-resize-handle.col-resize-left').nth(1);
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 102, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const after = await getColWidths(page);
    expect(after[0]).toBeGreaterThan(before[0] + 50);
    expect(after[1]).toBe(before[1]);
  });

  test('left-side handle on first data column resizes row-num column', async ({ page }) => {
    const beforeRowNum = await page.evaluate(() => app._test.windows[0].rowNumWidth || 50);
    const handle = page.locator('th:not(.row-num-header) .col-resize-handle.col-resize-left').first();
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 82, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const afterRowNum = await page.evaluate(() => app._test.windows[0].rowNumWidth || 50);
    expect(afterRowNum).toBeGreaterThan(beforeRowNum + 40);
  });

  test('resizing one column does not affect other columns', async ({ page }) => {
    const before = await getColWidths(page);
    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 102, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const after = await getColWidths(page);
    expect(after[1]).toBe(before[1]);
    expect(after[2]).toBe(before[2]);
  });

  test('resizing does not affect other columns in a maximized window', async ({ page }) => {
    await page.locator('.subwindow .btn-max').click();
    await page.waitForTimeout(300);

    const before = await getColWidths(page);
    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 102, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const after = await getColWidths(page);
    expect(after[0]).toBeGreaterThan(before[0] + 50);
    expect(after[1]).toBe(before[1]);
    expect(after[2]).toBe(before[2]);
  });

  test('column widths survive sort (rebuild)', async ({ page }) => {
    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 82, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const afterResize = await getColWidths(page);

    // Click header to sort (click center of th, away from resize handle)
    const th = page.locator('.subwindow table thead th').nth(1);
    const thBox = await th.boundingBox();
    await page.mouse.click(thBox.x + thBox.width / 3, thBox.y + thBox.height / 2);
    await page.waitForTimeout(300);

    const afterSort = await getColWidths(page);
    expect(afterSort).toEqual(afterResize);
  });

  test('column widths follow their column on reorder', async ({ page }) => {
    const before = await getColWidths(page);

    // Drag column 1 (email) to before column 0 (name)
    const fromTh = page.locator('.subwindow table thead th').nth(2);
    const toTh = page.locator('.subwindow table thead th').nth(1);
    const fromBox = await fromTh.boundingBox();
    const toBox = await toTh.boundingBox();

    await page.mouse.move(fromBox.x + fromBox.width / 2, fromBox.y + fromBox.height / 2);
    await page.mouse.down();
    await page.mouse.move(fromBox.x + fromBox.width / 2 - 15, fromBox.y + fromBox.height / 2, { steps: 5 });
    await page.mouse.move(toBox.x + toBox.width * 0.25, toBox.y + toBox.height / 2, { steps: 10 });
    await page.mouse.up();
    await page.waitForTimeout(300);

    const after = await getColWidths(page);
    // email was index 1, now index 0; name was index 0, now index 1
    expect(after[0]).toBe(before[1]);
    expect(after[1]).toBe(before[0]);
    expect(after[2]).toBe(before[2]);
  });

  test('minimum column width is enforced', async ({ page }) => {
    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x - 500, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const after = await getColWidths(page);
    expect(after[0]).toBeGreaterThanOrEqual(40);
  });

  test('double-click right-side resize handle auto-fits column width', async ({ page }) => {
    // First make column 0 very wide
    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 202, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const widened = await getColWidths(page);

    // Double-click the handle to auto-fit
    const box2 = await handle.boundingBox();
    await page.mouse.dblclick(box2.x + 2, box2.y + box2.height / 2);
    await page.waitForTimeout(300);

    const autoFit = await getColWidths(page);
    expect(autoFit[0]).toBeLessThan(widened[0]);
    expect(autoFit[0]).toBeGreaterThanOrEqual(40);
  });

  test('double-click left-side resize handle auto-fits the previous column', async ({ page }) => {
    // First make column 0 very wide via right-side handle
    const rightHandle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
    const rBox = await rightHandle.boundingBox();

    await page.mouse.move(rBox.x + 2, rBox.y + rBox.height / 2);
    await page.mouse.down();
    await page.mouse.move(rBox.x + 202, rBox.y + rBox.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const widened = await getColWidths(page);

    // Double-click left handle on column 1 to auto-fit column 0
    const leftHandle = page.locator('th:not(.row-num-header) .col-resize-handle.col-resize-left').nth(1);
    const lBox = await leftHandle.boundingBox();
    await page.mouse.dblclick(lBox.x + 2, lBox.y + lBox.height / 2);
    await page.waitForTimeout(300);

    const autoFit = await getColWidths(page);
    expect(autoFit[0]).toBeLessThan(widened[0]);
    expect(autoFit[0]).toBeGreaterThanOrEqual(40);
  });

  test('double-click left-side handle on first column auto-fits row-num column', async ({ page }) => {
    // Widen row-num column
    const rnHandle = page.locator('th.row-num-header .col-resize-handle');
    const rnBox = await rnHandle.boundingBox();
    await page.mouse.move(rnBox.x + 2, rnBox.y + rnBox.height / 2);
    await page.mouse.down();
    await page.mouse.move(rnBox.x + 152, rnBox.y + rnBox.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const widened = await page.evaluate(() => app._test.windows[0].rowNumWidth || 50);
    expect(widened).toBeGreaterThan(150);

    // Double-click left handle on first data column
    const leftHandle = page.locator('th:not(.row-num-header) .col-resize-handle.col-resize-left').first();
    const lBox = await leftHandle.boundingBox();
    await page.mouse.dblclick(lBox.x + 2, lBox.y + lBox.height / 2);
    await page.waitForTimeout(300);

    const autoFit = await page.evaluate(() => app._test.windows[0].rowNumWidth || 50);
    expect(autoFit).toBeLessThan(widened);
  });

  test('resize drag does not trigger column sort', async ({ page }) => {
    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 52, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(300);

    // Check that no sort was applied
    const sorted = await page.evaluate(() => {
      const win = app._test.windows[0];
      return win.sortCols.length;
    });
    expect(sorted).toBe(0);
  });

  test('left-side resize drag does not trigger column sort', async ({ page }) => {
    const handle = page.locator('th:not(.row-num-header) .col-resize-handle.col-resize-left').nth(1);
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 52, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(300);

    const sorted = await page.evaluate(() => {
      const win = app._test.windows[0];
      return win.sortCols.length;
    });
    expect(sorted).toBe(0);
  });

  test('colWidths reset when a column is added', async ({ page }) => {
    const before = await getColWidths(page);
    expect(before.length).toBe(3);

    // Right-click the # corner cell to insert a column
    const corner = page.locator('th.row-num-header');
    await corner.click({ button: 'right' });
    await page.locator('.context-menu button', { hasText: 'Insert Column' }).click();
    await page.locator('.modal-input').fill('newcol');
    await page.locator('.modal .ok').click();
    await page.waitForTimeout(500);

    // colWidths should have been re-measured with the new column
    const after = await getColWidths(page);
    expect(after).not.toBeNull();
    expect(after.length).toBe(4);
  });
});

function getRowNumWidth(page) {
  return page.evaluate(() => {
    const win = app._test.windows[0];
    return win.rowNumWidth || 50;
  });
}

test.describe('Row-number column resize', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('row-num header has a resize handle', async ({ page }) => {
    const handle = page.locator('th.row-num-header .col-resize-handle');
    await expect(handle).toHaveCount(1);
  });

  test('row-num resize handle has col-resize cursor', async ({ page }) => {
    const cursor = await page.locator('th.row-num-header .col-resize-handle').evaluate(
      el => getComputedStyle(el).cursor);
    expect(cursor).toBe('col-resize');
  });

  test('dragging row-num resize handle changes its width', async ({ page }) => {
    const before = await getRowNumWidth(page);
    const handle = page.locator('th.row-num-header .col-resize-handle');
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 82, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const after = await getRowNumWidth(page);
    expect(after).toBeGreaterThan(before + 40);
  });

  test('row-num resize does not affect data column widths', async ({ page }) => {
    const colsBefore = await getColWidths(page);
    const handle = page.locator('th.row-num-header .col-resize-handle');
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 82, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const colsAfter = await getColWidths(page);
    expect(colsAfter).toEqual(colsBefore);
  });

  test('row-num minimum width is enforced at 30px', async ({ page }) => {
    const handle = page.locator('th.row-num-header .col-resize-handle');
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x - 500, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const after = await getRowNumWidth(page);
    expect(after).toBe(30);
  });

  test('row-num col element width matches win.rowNumWidth', async ({ page }) => {
    const handle = page.locator('th.row-num-header .col-resize-handle');
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 62, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const { rowNumWidth, colElWidth } = await page.evaluate(() => {
      const win = app._test.windows[0];
      const colEl = win._colgroup.children[0];
      return {
        rowNumWidth: win.rowNumWidth,
        colElWidth: parseInt(colEl.style.width)
      };
    });
    expect(colElWidth).toBe(rowNumWidth);
  });

  test('double-click row-num resize handle auto-fits to row count', async ({ page }) => {
    // First make the column very wide
    const handle = page.locator('th.row-num-header .col-resize-handle');
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 152, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const widened = await getRowNumWidth(page);
    expect(widened).toBeGreaterThan(150);

    // Double-click to auto-fit
    const box2 = await handle.boundingBox();
    await page.mouse.dblclick(box2.x + 2, box2.y + box2.height / 2);
    await page.waitForTimeout(300);

    const autoFit = await getRowNumWidth(page);
    expect(autoFit).toBeLessThan(widened);
    expect(autoFit).toBeGreaterThanOrEqual(30);
  });

  test('row-num width survives sort (rebuild)', async ({ page }) => {
    const handle = page.locator('th.row-num-header .col-resize-handle');
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 62, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const afterResize = await getRowNumWidth(page);

    // Click a data column header to sort
    const th = page.locator('.subwindow table thead th').nth(1);
    const thBox = await th.boundingBox();
    await page.mouse.click(thBox.x + thBox.width / 3, thBox.y + thBox.height / 2);
    await page.waitForTimeout(300);

    const afterSort = await getRowNumWidth(page);
    expect(afterSort).toBe(afterResize);
  });

  test('row-num resize does not trigger select-all', async ({ page }) => {
    const handle = page.locator('th.row-num-header .col-resize-handle');
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 42, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const selectedCount = await page.evaluate(() => {
      const win = app._test.windows[0];
      return win.selectedCells ? win.selectedCells.size : 0;
    });
    expect(selectedCount).toBe(0);
  });

  test('row-num width not reset when a column is added', async ({ page }) => {
    // Resize row-num column
    const handle = page.locator('th.row-num-header .col-resize-handle');
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 62, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const beforeAdd = await getRowNumWidth(page);

    // Add a data column via corner cell context menu
    const corner = page.locator('th.row-num-header');
    await corner.click({ button: 'right' });
    await page.locator('.context-menu button', { hasText: 'Insert Column' }).click();
    await page.locator('.modal-input').fill('newcol');
    await page.locator('.modal .ok').click();
    await page.waitForTimeout(500);

    const afterAdd = await getRowNumWidth(page);
    expect(afterAdd).toBe(beforeAdd);
  });

  test('table width calculation includes row-num width', async ({ page }) => {
    const handle = page.locator('th.row-num-header .col-resize-handle');
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 102, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

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
});
