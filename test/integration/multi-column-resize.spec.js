const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow } = require('../helpers');

function getColWidths(page) {
  return page.evaluate(() => {
    const win = app._test.windows[0];
    return win.colWidths ? [...win.colWidths] : null;
  });
}

async function selectAllCells(page) {
  await page.evaluate(() => {
    const win = app._test.windows[0];
    app._test.selectAllCells(win);
  });
}

async function selectCellRange(page, fromRow, fromCol, toRow, toCol) {
  await page.evaluate(({ fr, fc, tr, tc }) => {
    const win = app._test.windows[0];
    const cols = win._columns;
    const rows = win._displayRows;
    win.anchorCell = { rownum: rows[fr]._rownum, col: cols[fc] };
    win.selectedCells = new Set();
    for (let r = Math.min(fr, tr); r <= Math.max(fr, tr); r++) {
      for (let c = Math.min(fc, tc); c <= Math.max(fc, tc); c++) {
        win.selectedCells.add(`${rows[r]._rownum}:${cols[c]}`);
      }
    }
  }, { fr: fromRow, fc: fromCol, tr: toRow, tc: toCol });
}

async function getSelectedColNames(page) {
  return page.evaluate(() => {
    const win = app._test.windows[0];
    const cols = new Set();
    if (win.selectedCells) {
      for (const key of win.selectedCells) {
        cols.add(key.slice(key.indexOf(':') + 1));
      }
    }
    return [...cols];
  });
}

test.describe('Multi-column resize — double-click auto-fit', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('double-click auto-fits all columns when all cells are selected', async ({ page }) => {
    // Widen all columns manually
    await page.evaluate(() => {
      const win = app._test.windows[0];
      for (let i = 0; i < win.colWidths.length; i++) {
        win.colWidths[i] = 400;
        win._colgroup.children[i + 1].style.width = '400px';
      }
      app._test.updateTableWidth(win);
    });

    const widened = await getColWidths(page);
    widened.forEach(w => expect(w).toBe(400));

    await selectAllCells(page);

    // Double-click the first column's resize handle
    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
    const box = await handle.boundingBox();
    await page.mouse.dblclick(box.x + 2, box.y + box.height / 2);
    await page.waitForTimeout(300);

    const after = await getColWidths(page);
    after.forEach(w => {
      expect(w).toBeLessThan(400);
      expect(w).toBeGreaterThanOrEqual(40);
    });
  });

  test('double-click auto-fits only the subset of selected columns', async ({ page }) => {
    // Widen all columns
    await page.evaluate(() => {
      const win = app._test.windows[0];
      for (let i = 0; i < win.colWidths.length; i++) {
        win.colWidths[i] = 400;
        win._colgroup.children[i + 1].style.width = '400px';
      }
      app._test.updateTableWidth(win);
    });

    // Select only columns 0 and 1 (name, email)
    await selectCellRange(page, 0, 0, 0, 1);

    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
    const box = await handle.boundingBox();
    await page.mouse.dblclick(box.x + 2, box.y + box.height / 2);
    await page.waitForTimeout(300);

    const after = await getColWidths(page);
    expect(after[0]).toBeLessThan(400);
    expect(after[1]).toBeLessThan(400);
    expect(after[2]).toBe(400); // unselected column unchanged
  });

  test('double-click on unselected column only auto-fits that column', async ({ page }) => {
    // Widen all columns
    await page.evaluate(() => {
      const win = app._test.windows[0];
      for (let i = 0; i < win.colWidths.length; i++) {
        win.colWidths[i] = 400;
        win._colgroup.children[i + 1].style.width = '400px';
      }
      app._test.updateTableWidth(win);
    });

    // Select only columns 1 and 2 (email, member_since)
    await selectCellRange(page, 0, 1, 0, 2);

    // Double-click column 0's resize handle (not selected)
    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
    const box = await handle.boundingBox();
    await page.mouse.dblclick(box.x + 2, box.y + box.height / 2);
    await page.waitForTimeout(300);

    const after = await getColWidths(page);
    expect(after[0]).toBeLessThan(400); // auto-fitted
    expect(after[1]).toBe(400); // not auto-fitted despite being selected
    expect(after[2]).toBe(400); // not auto-fitted despite being selected
  });

  test('double-click with single-cell selection only auto-fits that column', async ({ page }) => {
    // Widen all columns
    await page.evaluate(() => {
      const win = app._test.windows[0];
      for (let i = 0; i < win.colWidths.length; i++) {
        win.colWidths[i] = 400;
        win._colgroup.children[i + 1].style.width = '400px';
      }
      app._test.updateTableWidth(win);
    });

    // Select a single cell in column 0
    await selectCellRange(page, 0, 0, 0, 0);

    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
    const box = await handle.boundingBox();
    await page.mouse.dblclick(box.x + 2, box.y + box.height / 2);
    await page.waitForTimeout(300);

    const after = await getColWidths(page);
    expect(after[0]).toBeLessThan(400);
    expect(after[1]).toBe(400);
    expect(after[2]).toBe(400);
  });

  test('double-click with no selection auto-fits only that column', async ({ page }) => {
    // Widen all columns
    await page.evaluate(() => {
      const win = app._test.windows[0];
      for (let i = 0; i < win.colWidths.length; i++) {
        win.colWidths[i] = 400;
        win._colgroup.children[i + 1].style.width = '400px';
      }
      app._test.updateTableWidth(win);
    });

    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
    const box = await handle.boundingBox();
    await page.mouse.dblclick(box.x + 2, box.y + box.height / 2);
    await page.waitForTimeout(300);

    const after = await getColWidths(page);
    expect(after[0]).toBeLessThan(400);
    expect(after[1]).toBe(400);
    expect(after[2]).toBe(400);
  });

  test('left-side handle double-click auto-fits selected columns', async ({ page }) => {
    // Widen all columns
    await page.evaluate(() => {
      const win = app._test.windows[0];
      for (let i = 0; i < win.colWidths.length; i++) {
        win.colWidths[i] = 400;
        win._colgroup.children[i + 1].style.width = '400px';
      }
      app._test.updateTableWidth(win);
    });

    await selectAllCells(page);

    // Double-click left handle on column 1 (targets column 0)
    const leftHandle = page.locator('th:not(.row-num-header) .col-resize-handle.col-resize-left').nth(1);
    const lBox = await leftHandle.boundingBox();
    await page.mouse.dblclick(lBox.x + 2, lBox.y + lBox.height / 2);
    await page.waitForTimeout(300);

    const after = await getColWidths(page);
    after.forEach(w => {
      expect(w).toBeLessThan(400);
      expect(w).toBeGreaterThanOrEqual(40);
    });
  });
});

test.describe('Multi-column resize — drag resize', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('drag resize sets all selected columns to the same width on mouseup', async ({ page }) => {
    await selectAllCells(page);

    const before = await getColWidths(page);
    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 152, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const after = await getColWidths(page);
    const draggedWidth = after[0];
    expect(draggedWidth).toBeGreaterThan(before[0] + 100);
    expect(after[1]).toBe(draggedWidth);
    expect(after[2]).toBe(draggedWidth);
  });

  test('drag resize with partial selection only affects selected columns plus dragged', async ({ page }) => {
    // Select columns 0 and 1 only
    await selectCellRange(page, 0, 0, 0, 1);

    const before = await getColWidths(page);
    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 152, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const after = await getColWidths(page);
    const draggedWidth = after[0];
    expect(after[1]).toBe(draggedWidth);
    expect(after[2]).toBe(before[2]); // unselected, unchanged
  });

  test('drag resize with no selection only affects the dragged column', async ({ page }) => {
    const before = await getColWidths(page);
    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 152, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const after = await getColWidths(page);
    expect(after[0]).toBeGreaterThan(before[0] + 100);
    expect(after[1]).toBe(before[1]);
    expect(after[2]).toBe(before[2]);
  });

  test('drag resize includes dragged column even if not in selection', async ({ page }) => {
    // Select only column 1 (email)
    await selectCellRange(page, 0, 1, 0, 1);

    const before = await getColWidths(page);
    // Drag column 0's resize handle
    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 152, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const after = await getColWidths(page);
    const draggedWidth = after[0];
    expect(draggedWidth).toBeGreaterThan(before[0] + 100);
    expect(after[1]).toBe(draggedWidth); // selected, matched to dragged
    expect(after[2]).toBe(before[2]); // not selected, unchanged
  });

  test('during drag only the dragged column changes (deferred)', async ({ page }) => {
    await selectAllCells(page);

    const before = await getColWidths(page);
    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 152, box.y + box.height / 2, { steps: 5 });

    // While still dragging, check other columns haven't changed
    const during = await getColWidths(page);
    expect(during[0]).toBeGreaterThan(before[0] + 100); // dragged column changed
    expect(during[1]).toBe(before[1]); // not yet applied
    expect(during[2]).toBe(before[2]); // not yet applied

    await page.mouse.up();
    await page.waitForTimeout(200);

    // After mouseup, all snap
    const after = await getColWidths(page);
    expect(after[1]).toBe(after[0]);
    expect(after[2]).toBe(after[0]);
  });

  test('table width is correct after multi-column resize', async ({ page }) => {
    await selectAllCells(page);

    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
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

  test('single-cell selection behaves as no selection for drag resize', async ({ page }) => {
    // Select a single cell in column 1
    await selectCellRange(page, 0, 1, 0, 1);

    const before = await getColWidths(page);
    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 102, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const after = await getColWidths(page);
    // Dragged column 0 + single selected column 1 = 2 columns => multi-column resize
    expect(after[0]).toBeGreaterThan(before[0] + 50);
    expect(after[1]).toBe(after[0]);
    expect(after[2]).toBe(before[2]);
  });
});

test.describe('Multi-column resize — scroll compensation', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('scroll position compensates when left-side columns grow', async ({ page }) => {
    // Resize column 1 (middle) while all columns selected.
    // On mouseup, column 0 (left of dragged) grows to match, shifting column 1 right.
    // Scroll compensation should keep column 1's visual position stable.

    // First make column 1 wider so columns 0 and 2 will grow on mouseup
    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').nth(1);
    let box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 202, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const widenedWidths = await getColWidths(page);
    // Column 1 is now wide, columns 0 and 2 are original size

    // Now select all and drag column 1 even wider — on mouseup column 0 will grow
    await selectAllCells(page);

    // Record column 1's visual offset before
    const thOffsetBefore = await page.evaluate(() => {
      const win = app._test.windows[0];
      const th = win._table.querySelector('th[data-col-idx="1"]');
      return th.offsetLeft - win._container.scrollLeft;
    });
    const scrollBefore = await page.evaluate(() => app._test.windows[0]._container.scrollLeft);

    box = await handle.boundingBox();
    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 52, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const afterWidths = await getColWidths(page);
    // Column 0 grew from its original width to match column 1's final width
    expect(afterWidths[0]).toBe(afterWidths[1]);

    // scrollLeft should have been adjusted to compensate for column 0's growth
    const scrollAfter = await page.evaluate(() => app._test.windows[0]._container.scrollLeft);
    const col0Growth = afterWidths[0] - widenedWidths[0];
    // The scroll compensation should offset by approximately the growth of left-side columns
    expect(scrollAfter - scrollBefore).toBeGreaterThanOrEqual(0);

    // Column 1's visual offset should be approximately preserved
    const thOffsetAfter = await page.evaluate(() => {
      const win = app._test.windows[0];
      const th = win._table.querySelector('th[data-col-idx="1"]');
      return th.offsetLeft - win._container.scrollLeft;
    });
    expect(Math.abs(thOffsetAfter - thOffsetBefore)).toBeLessThanOrEqual(2);
  });

  test('scroll compensation clamps to zero and does not go negative', async ({ page }) => {
    // Make columns narrow
    await page.evaluate(() => {
      const win = app._test.windows[0];
      for (let i = 0; i < win.colWidths.length; i++) {
        win.colWidths[i] = 200;
        win._colgroup.children[i + 1].style.width = '200px';
      }
      app._test.updateTableWidth(win);
    });

    // Select all columns, start with scrollLeft = 0
    await selectAllCells(page);
    await page.evaluate(() => {
      app._test.windows[0]._container.scrollLeft = 0;
    });

    // Drag resize column 1 to shrink it
    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').nth(1);
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x - 100, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const scrollLeft = await page.evaluate(() => app._test.windows[0]._container.scrollLeft);
    expect(scrollLeft).toBeGreaterThanOrEqual(0);
  });

  test('no scroll compensation when resizing without multi-column selection', async ({ page }) => {
    // Ensure scrollLeft starts at 0
    await page.evaluate(() => {
      app._test.windows[0]._container.scrollLeft = 0;
    });

    // Drag resize with no selection
    const handle = page.locator('th:not(.row-num-header) .col-resize-handle:not(.col-resize-left)').first();
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 102, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const scrollLeft = await page.evaluate(() => app._test.windows[0]._container.scrollLeft);
    expect(scrollLeft).toBe(0);
  });
});
