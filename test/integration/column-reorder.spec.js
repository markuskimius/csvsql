const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, getTableData, executeSQL } = require('../helpers');

// Find the window object for a given table name via app._test.
async function getWin(page, tableName) {
  return page.evaluate((name) => {
    const w = app._test.windows.find(w => w.tableName === name);
    return w ? { id: w.id, selectedCol: w.selectedCol, sortCols: w.sortCols } : null;
  }, tableName);
}

async function getColumns(page, tableName) {
  const t = await getTableData(page, tableName);
  return t.columns;
}

test.describe('Column reorder', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('clean click on header selects the column without sorting', async ({ page }) => {
    const cols = await getColumns(page, 'sample1');
    const firstHeader = page.locator('.subwindow table thead th').nth(1); // [0] is row-num
    await firstHeader.click();

    await expect(firstHeader).toHaveClass(/col-highlight/);

    const win = await getWin(page, 'sample1');
    expect(win.selectedCol).toBe(cols[0]);
    expect(win.sortCols.length).toBe(0);
  });

  test('clicking sort badge toggles sort', async ({ page }) => {
    const firstHeader = page.locator('.subwindow table thead th').nth(1);
    const sortBadge = firstHeader.locator('.col-badge-sort');
    await sortBadge.click();
    await page.waitForTimeout(200);

    await expect(firstHeader).toHaveClass(/sorted/);
    const win = await getWin(page, 'sample1');
    expect(win.sortCols.length).toBe(1);
    expect(win.sortCols[0].dir).toBe('asc');
  });

  test('plain drag reorders columns', async ({ page }) => {
    const before = await getColumns(page, 'sample1');
    expect(before).toEqual(['name', 'email', 'member_since']);

    // Drag email header leftward past midpoint of name header (no modifier).
    const fromTh = page.locator('.subwindow table thead th').nth(2); // email
    const toTh = page.locator('.subwindow table thead th').nth(1);   // name

    const fromBox = await fromTh.boundingBox();
    const toBox = await toTh.boundingBox();

    await page.mouse.move(fromBox.x + fromBox.width / 2, fromBox.y + fromBox.height / 2);
    await page.mouse.down();
    await page.mouse.move(fromBox.x + fromBox.width / 2 - 15, fromBox.y + fromBox.height / 2, { steps: 5 });
    await page.mouse.move(toBox.x + toBox.width * 0.25, toBox.y + toBox.height / 2, { steps: 10 });
    await page.mouse.up();

    await page.waitForTimeout(300);
    const after = await getColumns(page, 'sample1');
    expect(after).toEqual(['email', 'name', 'member_since']);
  });

  test('Ctrl+Right moves selected column right; Ctrl+Left moves it back', async ({ page }) => {
    // Select first column
    const firstHeader = page.locator('.subwindow table thead th').nth(1);
    await firstHeader.click();

    const before = await getColumns(page, 'sample1');
    expect(before).toEqual(['name', 'email', 'member_since']);

    // Focus document body so keydown isn't captured by an input
    await page.evaluate(() => document.body.focus());
    await page.keyboard.press('Control+ArrowRight');
    await page.waitForTimeout(200);

    const afterRight = await getColumns(page, 'sample1');
    expect(afterRight).toEqual(['email', 'name', 'member_since']);

    // Selection should still be on "name" (now at index 1)
    const win = await getWin(page, 'sample1');
    expect(win.selectedCol).toBe('name');

    await page.keyboard.press('Control+ArrowLeft');
    await page.waitForTimeout(200);

    const afterLeft = await getColumns(page, 'sample1');
    expect(afterLeft).toEqual(['name', 'email', 'member_since']);
  });

  test('Ctrl+Left on leftmost column is a no-op', async ({ page }) => {
    const firstHeader = page.locator('.subwindow table thead th').nth(1);
    await firstHeader.click();

    await page.evaluate(() => document.body.focus());
    await page.keyboard.press('Control+ArrowLeft');
    await page.waitForTimeout(200);

    const after = await getColumns(page, 'sample1');
    expect(after).toEqual(['name', 'email', 'member_since']);
  });

  test('Ctrl+Arrow with no selection is a no-op', async ({ page }) => {
    const before = await getColumns(page, 'sample1');

    await page.evaluate(() => document.body.focus());
    await page.keyboard.press('Control+ArrowRight');
    await page.waitForTimeout(200);

    const after = await getColumns(page, 'sample1');
    expect(after).toEqual(before);
  });

  test('Ctrl+Arrow is ignored when focus is in an input (SQL console)', async ({ page }) => {
    // Select a column first
    const firstHeader = page.locator('.subwindow table thead th').nth(1);
    await firstHeader.click();

    // Put focus in the SQL input and press Ctrl+Right
    await page.locator('#sql-input').focus();
    await page.keyboard.press('Control+ArrowRight');
    await page.waitForTimeout(200);

    const after = await getColumns(page, 'sample1');
    expect(after).toEqual(['name', 'email', 'member_since']);
  });

  test('multiple Ctrl+Right presses walk the column to the end', async ({ page }) => {
    // Create a 3-column query result to test middle → end traversal.
    await executeSQL(page, "SELECT 'x' AS a, 'y' AS b, 'z' AS c");
    // The new window's table name follows the q_ pattern; grab it.
    const qName = await page.evaluate(() => {
      const w = app._test.windows.find(w => w.isQuery);
      return w ? w.tableName : null;
    });
    expect(qName).toBeTruthy();

    // Focus that query window and select its first column
    await page.evaluate((name) => {
      const w = app._test.windows.find(w => w.tableName === name);
      w.el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    }, qName);

    // Click the header's label (auto-fit columns are narrow, so the th
    // center could land on the sort/filter buttons instead)
    const firstHeader = page.locator(`.subwindow.active table thead th`).nth(1);
    await firstHeader.locator('.col-name').click();

    await page.evaluate(() => document.body.focus());
    await page.keyboard.press('Control+ArrowRight');
    await page.waitForTimeout(150);
    await page.keyboard.press('Control+ArrowRight');
    await page.waitForTimeout(150);

    const cols = await getColumns(page, qName);
    expect(cols).toEqual(['b', 'c', 'a']);

    // One more Right at the right edge: no-op
    await page.keyboard.press('Control+ArrowRight');
    await page.waitForTimeout(150);
    const after = await getColumns(page, qName);
    expect(after).toEqual(['b', 'c', 'a']);
  });

  test('drag column to the last position', async ({ page }) => {
    const before = await getColumns(page, 'sample1');
    expect(before).toEqual(['name', 'email', 'member_since']);

    // Drag "name" (first col) past the right half of "member_since" (last col).
    const fromTh = page.locator('.subwindow table thead th').nth(1); // name
    const toTh = page.locator('.subwindow table thead th').nth(3);   // member_since

    const fromBox = await fromTh.boundingBox();
    const toBox = await toTh.boundingBox();

    await page.mouse.move(fromBox.x + fromBox.width / 2, fromBox.y + fromBox.height / 2);
    await page.mouse.down();
    await page.mouse.move(fromBox.x + fromBox.width / 2 + 15, fromBox.y + fromBox.height / 2, { steps: 3 });
    await page.mouse.move(toBox.x + toBox.width * 0.75, toBox.y + toBox.height / 2, { steps: 10 });
    await page.mouse.up();

    await page.waitForTimeout(300);
    const after = await getColumns(page, 'sample1');
    expect(after).toEqual(['email', 'member_since', 'name']);
  });

  test('drag column from last to first position', async ({ page }) => {
    const before = await getColumns(page, 'sample1');
    expect(before).toEqual(['name', 'email', 'member_since']);

    // Drag "member_since" (last col) to the left half of "name" (first col).
    const fromTh = page.locator('.subwindow table thead th').nth(3); // member_since
    const toTh = page.locator('.subwindow table thead th').nth(1);   // name

    const fromBox = await fromTh.boundingBox();
    const toBox = await toTh.boundingBox();

    await page.mouse.move(fromBox.x + fromBox.width / 2, fromBox.y + fromBox.height / 2);
    await page.mouse.down();
    await page.mouse.move(fromBox.x + fromBox.width / 2 - 15, fromBox.y + fromBox.height / 2, { steps: 3 });
    await page.mouse.move(toBox.x + toBox.width * 0.25, toBox.y + toBox.height / 2, { steps: 10 });
    await page.mouse.up();

    await page.waitForTimeout(300);
    const after = await getColumns(page, 'sample1');
    expect(after).toEqual(['member_since', 'name', 'email']);
  });

  test('drag middle column to last position', async ({ page }) => {
    const before = await getColumns(page, 'sample1');
    expect(before).toEqual(['name', 'email', 'member_since']);

    // Drag "email" (middle col) past the right half of "member_since" (last col).
    const fromTh = page.locator('.subwindow table thead th').nth(2); // email
    const toTh = page.locator('.subwindow table thead th').nth(3);   // member_since

    const fromBox = await fromTh.boundingBox();
    const toBox = await toTh.boundingBox();

    await page.mouse.move(fromBox.x + fromBox.width / 2, fromBox.y + fromBox.height / 2);
    await page.mouse.down();
    await page.mouse.move(fromBox.x + fromBox.width / 2 + 15, fromBox.y + fromBox.height / 2, { steps: 3 });
    await page.mouse.move(toBox.x + toBox.width * 0.75, toBox.y + toBox.height / 2, { steps: 10 });
    await page.mouse.up();

    await page.waitForTimeout(300);
    const after = await getColumns(page, 'sample1');
    expect(after).toEqual(['name', 'member_since', 'email']);
  });

  test('sorting works on large numeric values that exceed Number precision', async ({ page }) => {
    // member_since has values like 1780862826.123456700 that differ only
    // in the last digits — beyond Number.MAX_SAFE_INTEGER when combined.
    // Sort should still differentiate them via string comparison fallback.
    const memberSinceHeader = page.locator('.subwindow table thead th').nth(3); // member_since
    const sortBadge = memberSinceHeader.locator('.col-badge-sort');
    await sortBadge.click();
    await page.waitForTimeout(200);

    // After ascending sort, verify the rows are in order
    const values = await page.evaluate(() => {
      const win = app._test.windows[0];
      return win._displayRows.map(r => r.member_since);
    });

    for (let i = 1; i < values.length; i++) {
      expect(String(values[i]) >= String(values[i - 1])).toBe(true);
    }

    // Click again for descending
    await sortBadge.click();
    await page.waitForTimeout(200);

    const desc = await page.evaluate(() => {
      const win = app._test.windows[0];
      return win._displayRows.map(r => r.member_since);
    });

    for (let i = 1; i < desc.length; i++) {
      expect(String(desc[i]) <= String(desc[i - 1])).toBe(true);
    }
  });
});
