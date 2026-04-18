const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, getTableData } = require('../helpers');

// Dispatch a touch event on an element at the given viewport coordinates.
async function dispatchTouch(page, elementHandle, type, clientX, clientY) {
  await elementHandle.evaluate((el, args) => {
    const touch = new Touch({
      identifier: 1,
      target: el,
      clientX: args.x,
      clientY: args.y,
      pageX: args.x,
      pageY: args.y,
      radiusX: 1,
      radiusY: 1,
      force: 1,
    });
    const list = args.type === 'touchend' || args.type === 'touchcancel' ? [] : [touch];
    const evt = new TouchEvent(args.type, {
      cancelable: true,
      bubbles: true,
      touches: list,
      targetTouches: list,
      changedTouches: [touch],
    });
    el.dispatchEvent(evt);
  }, { type, x: clientX, y: clientY });
}

// Dispatch a pointer event on an element at the given viewport coordinates.
async function dispatchPointer(page, elementHandle, type, clientX, clientY, pointerType = 'touch') {
  await elementHandle.evaluate((el, args) => {
    const evt = new PointerEvent(args.type, {
      cancelable: true,
      bubbles: true,
      pointerType: args.pointerType,
      isPrimary: true,
      pointerId: 1,
      clientX: args.x,
      clientY: args.y,
    });
    el.dispatchEvent(evt);
  }, { type, x: clientX, y: clientY, pointerType });
}

async function centerOf(locator) {
  const box = await locator.boundingBox();
  return { x: box.x + box.width / 2, y: box.y + box.height / 2 };
}

test.describe('Touch gestures', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('long-press on a cell enters edit mode', async ({ page }) => {
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    const handle = await cell.elementHandle();
    const { x, y } = await centerOf(cell);

    await dispatchTouch(page, handle, 'touchstart', x, y);
    // Hold past LONG_PRESS_MS (600ms) — use extra margin for CI timing
    await page.waitForTimeout(750);
    await dispatchTouch(page, handle, 'touchend', x, y);

    await expect(cell).toHaveAttribute('contenteditable', 'true');
  });

  test('short tap on a cell does not enter edit mode', async ({ page }) => {
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    const handle = await cell.elementHandle();
    const { x, y } = await centerOf(cell);

    await dispatchTouch(page, handle, 'touchstart', x, y);
    await page.waitForTimeout(100);
    await dispatchTouch(page, handle, 'touchend', x, y);
    await page.waitForTimeout(200);

    await expect(cell).not.toHaveAttribute('contenteditable', 'true');
  });

  test('long-press via pointer events also enters edit mode', async ({ page }) => {
    // Some browsers / devtools modes deliver pointer events but not touch events;
    // the handler should work in either path.
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    const handle = await cell.elementHandle();
    const { x, y } = await centerOf(cell);

    await dispatchPointer(page, handle, 'pointerdown', x, y);
    await page.waitForTimeout(750);
    await dispatchPointer(page, handle, 'pointerup', x, y);

    await expect(cell).toHaveAttribute('contenteditable', 'true');
  });

  test('pointer long-press ignores mouse input', async ({ page }) => {
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    const handle = await cell.elementHandle();
    const { x, y } = await centerOf(cell);

    await dispatchPointer(page, handle, 'pointerdown', x, y, 'mouse');
    await page.waitForTimeout(750);
    await dispatchPointer(page, handle, 'pointerup', x, y, 'mouse');

    await expect(cell).not.toHaveAttribute('contenteditable', 'true');
  });

  test('long-press cancelled by significant finger movement', async ({ page }) => {
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    const handle = await cell.elementHandle();
    const { x, y } = await centerOf(cell);

    await dispatchTouch(page, handle, 'touchstart', x, y);
    await page.waitForTimeout(200);
    await dispatchTouch(page, handle, 'touchmove', x + 30, y);
    await page.waitForTimeout(500);
    await dispatchTouch(page, handle, 'touchend', x + 30, y);

    await expect(cell).not.toHaveAttribute('contenteditable', 'true');
  });

  test('1.5-tap reorders columns', async ({ page }) => {
    const before = (await getTableData(page, 'sample1')).columns;
    expect(before).toEqual(['name', 'email']);

    const nameTh = page.locator('.subwindow table thead th').nth(1);  // name
    const emailTh = page.locator('.subwindow table thead th').nth(2); // email

    // Tap 1: single tap on email header (the column we want to move).
    {
      const h = await emailTh.elementHandle();
      const c = await centerOf(emailTh);
      await dispatchTouch(page, h, 'touchstart', c.x, c.y);
      await page.waitForTimeout(50);
      await dispatchTouch(page, h, 'touchend', c.x, c.y);
      await emailTh.dispatchEvent('click');
    }

    // Tap 2: press email and pan leftward into the left half of name.
    await page.waitForTimeout(100);
    {
      const h = await emailTh.elementHandle();
      const start = await centerOf(emailTh);
      const nameBox = await nameTh.boundingBox();
      const dropX = nameBox.x + nameBox.width * 0.25; // left half of name
      const dropY = nameBox.y + nameBox.height / 2;

      await dispatchTouch(page, h, 'touchstart', start.x, start.y);
      await dispatchTouch(page, h, 'touchmove', start.x - 15, start.y);
      await dispatchTouch(page, h, 'touchmove', dropX, dropY);
      await dispatchTouch(page, h, 'touchend', dropX, dropY);
    }

    await page.waitForTimeout(300);
    const after = (await getTableData(page, 'sample1')).columns;
    expect(after).toEqual(['email', 'name']);
  });

  test('1.5-tap on a cell extends the selection rectangle', async ({ page }) => {
    // Two rows × two cols of data; pick a start and end cell to define a rect.
    const startCell = page.locator('.subwindow table tbody tr').nth(0).locator('td.data-cell').nth(0);
    const endCell = page.locator('.subwindow table tbody tr').nth(2).locator('td.data-cell').nth(1);

    const startBox = await startCell.boundingBox();
    const startCtr = { x: startBox.x + startBox.width / 2, y: startBox.y + startBox.height / 2 };
    const endBox = await endCell.boundingBox();
    const endCtr = { x: endBox.x + endBox.width / 2, y: endBox.y + endBox.height / 2 };

    // Tap 1: brief tap on the anchor cell.
    {
      const h = await startCell.elementHandle();
      await dispatchTouch(page, h, 'touchstart', startCtr.x, startCtr.y);
      await page.waitForTimeout(50);
      await dispatchTouch(page, h, 'touchend', startCtr.x, startCtr.y);
      await startCell.dispatchEvent('click');
    }

    // Tap 2: tap-and-pan across to a different cell to build the rectangle.
    await page.waitForTimeout(100);
    {
      const h = await startCell.elementHandle();
      await dispatchTouch(page, h, 'touchstart', startCtr.x, startCtr.y);
      await dispatchTouch(page, h, 'touchmove', startCtr.x + 15, startCtr.y);
      await dispatchTouch(page, h, 'touchmove', endCtr.x, endCtr.y);
      await dispatchTouch(page, h, 'touchend', endCtr.x, endCtr.y);
    }

    await page.waitForTimeout(200);
    const selected = await page.evaluate(() => {
      const win = app._test.windows[0];
      return win.selectedCells ? [...win.selectedCells].sort() : [];
    });
    // 2 cols × 3 rows = 6 cells
    expect(selected.length).toBe(6);
    expect(selected).toContain('1:name');
    expect(selected).toContain('1:email');
    expect(selected).toContain('2:name');
    expect(selected).toContain('2:email');
    expect(selected).toContain('3:name');
    expect(selected).toContain('3:email');
  });

  test('1.5-tap on titlebar moves the window', async ({ page }) => {
    const titlebar = page.locator('.subwindow .win-titlebar').first();
    const before = await page.evaluate(() => {
      const el = document.querySelector('.subwindow');
      return { left: parseInt(el.style.left), top: parseInt(el.style.top) };
    });
    const box = await titlebar.boundingBox();
    const startX = box.x + 30;
    const startY = box.y + box.height / 2;

    const h = await titlebar.elementHandle();
    // Tap 1: brief tap
    await dispatchTouch(page, h, 'touchstart', startX, startY);
    await page.waitForTimeout(50);
    await dispatchTouch(page, h, 'touchend', startX, startY);

    // Tap 2: tap-and-pan
    await page.waitForTimeout(100);
    await dispatchTouch(page, h, 'touchstart', startX, startY);
    await dispatchTouch(page, h, 'touchmove', startX + 20, startY);
    await dispatchTouch(page, h, 'touchmove', startX + 80, startY + 40);
    await dispatchTouch(page, h, 'touchend', startX + 80, startY + 40);

    await page.waitForTimeout(150);
    const after = await page.evaluate(() => {
      const el = document.querySelector('.subwindow');
      return { left: parseInt(el.style.left), top: parseInt(el.style.top) };
    });
    expect(after.left).toBeGreaterThan(before.left + 40);
    expect(after.top).toBeGreaterThan(before.top + 20);
  });

  test('single touch on titlebar does not move the window', async ({ page }) => {
    const titlebar = page.locator('.subwindow .win-titlebar').first();
    const before = await page.evaluate(() => {
      const el = document.querySelector('.subwindow');
      return { left: parseInt(el.style.left), top: parseInt(el.style.top) };
    });
    const box = await titlebar.boundingBox();
    const startX = box.x + 30;
    const startY = box.y + box.height / 2;

    const h = await titlebar.elementHandle();
    // Single tap-and-pan — no prior tap, so should NOT move the window.
    await dispatchTouch(page, h, 'touchstart', startX, startY);
    await dispatchTouch(page, h, 'touchmove', startX + 80, startY + 40);
    await dispatchTouch(page, h, 'touchend', startX + 80, startY + 40);

    await page.waitForTimeout(150);
    const after = await page.evaluate(() => {
      const el = document.querySelector('.subwindow');
      return { left: parseInt(el.style.left), top: parseInt(el.style.top) };
    });
    expect(after.left).toBe(before.left);
    expect(after.top).toBe(before.top);
  });

  test('tap without follow-up does not enter drag mode', async ({ page }) => {
    const before = (await getTableData(page, 'sample1')).columns;
    expect(before).toEqual(['name', 'email']);

    const nameTh = page.locator('.subwindow table thead th').nth(1);
    const handle = await nameTh.elementHandle();
    const { x, y } = await centerOf(nameTh);

    // Single tap — no second tap-and-pan
    await dispatchTouch(page, handle, 'touchstart', x, y);
    await page.waitForTimeout(50);
    await dispatchTouch(page, handle, 'touchend', x, y);
    await page.waitForTimeout(800); // past the double-tap window

    // Second tap now — but outside the double-tap window, should not start drag
    const emailTh = page.locator('.subwindow table thead th').nth(2);
    const emailCtr = await centerOf(emailTh);
    const handle2 = await nameTh.elementHandle();
    const nameCtr = await centerOf(nameTh);
    await dispatchTouch(page, handle2, 'touchstart', nameCtr.x, nameCtr.y);
    await dispatchTouch(page, handle2, 'touchmove', emailCtr.x + 5, emailCtr.y);
    await dispatchTouch(page, handle2, 'touchend', emailCtr.x + 5, emailCtr.y);
    await page.waitForTimeout(300);

    const after = (await getTableData(page, 'sample1')).columns;
    expect(after).toEqual(['name', 'email']);
  });
});
