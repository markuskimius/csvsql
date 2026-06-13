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

  test('resize handles exist on each column header', async ({ page }) => {
    const handleCount = await page.locator('.col-resize-handle').count();
    const colCount = await page.evaluate(() => app._test.tables.sample1.columns.length);
    expect(handleCount).toBe(colCount);
  });

  test('resize handle has col-resize cursor', async ({ page }) => {
    const cursor = await page.locator('.col-resize-handle').first().evaluate(
      el => getComputedStyle(el).cursor);
    expect(cursor).toBe('col-resize');
  });

  test('dragging a resize handle changes that column width', async ({ page }) => {
    const before = await getColWidths(page);
    const handle = page.locator('.col-resize-handle').first();
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 102, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const after = await getColWidths(page);
    expect(after[0]).toBeGreaterThan(before[0] + 50);
  });

  test('resizing one column does not affect other columns', async ({ page }) => {
    const before = await getColWidths(page);
    const handle = page.locator('.col-resize-handle').first();
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
    const handle = page.locator('.col-resize-handle').first();
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
    const handle = page.locator('.col-resize-handle').first();
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
    const handle = page.locator('.col-resize-handle').first();
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x - 500, box.y + box.height / 2, { steps: 5 });
    await page.mouse.up();
    await page.waitForTimeout(200);

    const after = await getColWidths(page);
    expect(after[0]).toBeGreaterThanOrEqual(40);
  });

  test('double-click resize handle auto-fits column width', async ({ page }) => {
    // First make column 0 very wide
    const handle = page.locator('.col-resize-handle').first();
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

  test('resize drag does not trigger column sort', async ({ page }) => {
    const handle = page.locator('.col-resize-handle').first();
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

  test('colWidths reset when a column is added', async ({ page }) => {
    const before = await getColWidths(page);
    expect(before.length).toBe(3);

    // Click + Col to open the custom modal prompt
    await page.locator('.win-toolbar button', { hasText: '+ Col' }).click();
    await page.locator('.modal-input').fill('newcol');
    await page.locator('.modal .ok').click();
    await page.waitForTimeout(500);

    // colWidths should have been re-measured with the new column
    const after = await getColWidths(page);
    expect(after).not.toBeNull();
    expect(after.length).toBe(4);
  });
});
