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
