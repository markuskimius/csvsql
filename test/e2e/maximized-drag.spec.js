const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, executeSQL } = require('../helpers');

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

async function getWinState(page) {
  return page.evaluate(() => {
    const w = app._test.windows[0];
    return {
      maximized: w.maximized,
      left: parseInt(w.el.style.left),
      top: parseInt(w.el.style.top),
      width: w.el.offsetWidth,
      height: w.el.offsetHeight,
    };
  });
}

test('dragging titlebar of maximized window restores size and moves it', async ({ page }) => {
  await openApp(page);
  await uploadFile(page, 'sample1.csv');
  await waitForWindow(page, 'sample1');

  const before = await getWinState(page);

  await page.evaluate(() => app._test.toggleMaximize(app._test.windows[0].id));
  const maxState = await getWinState(page);
  expect(maxState.maximized).toBe(true);
  expect(maxState.width).toBeGreaterThan(before.width);

  const titlebar = page.locator('.subwindow .win-titlebar').first();
  const box = await titlebar.boundingBox();
  const grabX = box.x + box.width / 2;
  const grabY = box.y + box.height / 2;
  await page.mouse.move(grabX, grabY);
  await page.mouse.down();
  await page.mouse.move(grabX + 60, grabY + 80, { steps: 5 });
  await page.mouse.up();

  const after = await getWinState(page);
  expect(after.maximized).toBe(false);
  // Size restored to pre-maximize bounds
  expect(after.width).toBe(before.width);
  expect(after.height).toBe(before.height);
  // Window moved with the drag (not stuck at 0,0)
  expect(after.top).toBeGreaterThan(maxState.top);
});

test('grab point stays proportionally on the titlebar after restore', async ({ page }) => {
  await openApp(page);
  await uploadFile(page, 'sample1.csv');
  await waitForWindow(page, 'sample1');

  const before = await getWinState(page);
  await page.evaluate(() => app._test.toggleMaximize(app._test.windows[0].id));

  const area = await page.locator('#window-area').boundingBox();
  // Grab near the right edge of the maximized titlebar (90% across),
  // avoiding the window buttons at the far right.
  const titlebar = page.locator('.subwindow .win-titlebar').first();
  const box = await titlebar.boundingBox();
  const grabX = box.x + box.width * 0.8;
  const grabY = box.y + box.height / 2;
  await page.mouse.move(grabX, grabY);
  await page.mouse.down();
  await page.mouse.move(grabX, grabY + 50, { steps: 5 });

  const during = await getWinState(page);
  // Restored window should be positioned so the cursor is ~80% across it
  const cursorOffset = (grabX - area.x) - during.left;
  const expected = before.width * 0.8;
  expect(Math.abs(cursorOffset - expected)).toBeLessThan(15);
  await page.mouse.up();
});

test('click without drag keeps window maximized', async ({ page }) => {
  await openApp(page);
  await uploadFile(page, 'sample1.csv');
  await waitForWindow(page, 'sample1');

  await page.evaluate(() => app._test.toggleMaximize(app._test.windows[0].id));
  const maxState = await getWinState(page);

  const titlebar = page.locator('.subwindow .win-titlebar').first();
  const box = await titlebar.boundingBox();
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  await page.mouse.down();
  await page.mouse.up();

  const after = await getWinState(page);
  expect(after.maximized).toBe(true);
  expect(after.width).toBe(maxState.width);
});

test('dragging tab bar of maximized dock restores size and moves it', async ({ page }) => {
  await openApp(page);
  await uploadFile(page, 'sample1.csv');
  await waitForWindow(page, 'sample1');
  await executeSQL(page, 'SELECT * INTO [t2] FROM [sample1]');
  await waitForWindow(page, 't2');

  const preDock = await page.evaluate(() => {
    const w = app._test.windows;
    app._test.mergeWindowsIntoTabs(w[0].id, w[1].id);
    const dock = app._test.dockContainers[0];
    return { width: dock.el.offsetWidth, height: dock.el.offsetHeight };
  });

  await page.evaluate(() => app._test.toggleMaximizeDock(app._test.dockContainers[0]));
  const maxDock = await page.evaluate(() => {
    const dock = app._test.dockContainers[0];
    return { maximized: dock.maximized, width: dock.el.offsetWidth };
  });
  expect(maxDock.maximized).toBe(true);

  // Drag the empty area of the tab bar just right of the last tab
  // (the far right edge is covered by the dock control buttons)
  const lastTab = page.locator('.dock-tab').last();
  const box = await lastTab.boundingBox();
  const grabX = box.x + box.width + 15;
  const grabY = box.y + box.height / 2;
  await page.mouse.move(grabX, grabY);
  await page.mouse.down();
  await page.mouse.move(grabX - 40, grabY + 60, { steps: 5 });
  await page.mouse.up();

  const after = await page.evaluate(() => {
    const dock = app._test.dockContainers[0];
    return {
      maximized: dock.maximized,
      top: parseInt(dock.el.style.top),
      width: dock.el.offsetWidth,
      height: dock.el.offsetHeight,
    };
  });
  expect(after.maximized).toBe(false);
  expect(after.width).toBe(preDock.width);
  expect(after.height).toBe(preDock.height);
  expect(after.top).toBeGreaterThan(0);
});

test('grab near left edge keeps window under cursor at ratio ~0', async ({ page }) => {
  await openApp(page);
  await uploadFile(page, 'sample1.csv');
  await waitForWindow(page, 'sample1');

  await page.evaluate(() => app._test.toggleMaximize(app._test.windows[0].id));

  const area = await page.locator('#window-area').boundingBox();
  const titlebar = page.locator('.subwindow .win-titlebar').first();
  const box = await titlebar.boundingBox();
  // Grab near the left edge (5% across the maximized titlebar)
  const grabX = box.x + box.width * 0.05;
  const grabY = box.y + box.height / 2;
  await page.mouse.move(grabX, grabY);
  await page.mouse.down();
  await page.mouse.move(grabX, grabY + 50, { steps: 5 });

  const during = await getWinState(page);
  // Window's left edge should sit close to the cursor (small ratio offset)
  const cursorOffset = (grabX - area.x) - during.left;
  expect(cursorOffset).toBeGreaterThanOrEqual(0);
  expect(cursorOffset).toBeLessThan(during.width * 0.1 + 15);
  // And never outside the workspace
  expect(during.left).toBeGreaterThanOrEqual(0);
  await page.mouse.up();
});

test('re-maximize after drag-restore captures the new position', async ({ page }) => {
  await openApp(page);
  await uploadFile(page, 'sample1.csv');
  await waitForWindow(page, 'sample1');

  const before = await getWinState(page);
  await page.evaluate(() => app._test.toggleMaximize(app._test.windows[0].id));

  const titlebar = page.locator('.subwindow .win-titlebar').first();
  const box = await titlebar.boundingBox();
  const grabX = box.x + box.width / 2;
  const grabY = box.y + box.height / 2;
  await page.mouse.move(grabX, grabY);
  await page.mouse.down();
  await page.mouse.move(grabX + 40, grabY + 90, { steps: 5 });
  await page.mouse.up();

  const dragged = await getWinState(page);
  expect(dragged.maximized).toBe(false);
  expect(dragged.width).toBe(before.width);

  // Maximize again, then restore: should return to the dragged position
  await page.evaluate(() => app._test.toggleMaximize(app._test.windows[0].id));
  const maxAgain = await getWinState(page);
  expect(maxAgain.maximized).toBe(true);
  await page.evaluate(() => app._test.toggleMaximize(app._test.windows[0].id));
  const restored = await getWinState(page);
  expect(restored.maximized).toBe(false);
  expect(restored.left).toBe(dragged.left);
  expect(restored.top).toBe(dragged.top);
  expect(restored.width).toBe(dragged.width);
  expect(restored.height).toBe(dragged.height);
});

test('touch drag (1.5-tap) on maximized titlebar restores size and moves it', async ({ page }) => {
  await openApp(page);
  await uploadFile(page, 'sample1.csv');
  await waitForWindow(page, 'sample1');

  const before = await getWinState(page);
  await page.evaluate(() => app._test.toggleMaximize(app._test.windows[0].id));
  const maxState = await getWinState(page);
  expect(maxState.maximized).toBe(true);

  const titlebar = page.locator('.subwindow .win-titlebar').first();
  const handle = await titlebar.elementHandle();
  const box = await titlebar.boundingBox();
  const x = box.x + box.width / 2;
  const y = box.y + box.height / 2;

  // Tap 1
  await dispatchTouch(page, handle, 'touchstart', x, y);
  await page.waitForTimeout(50);
  await dispatchTouch(page, handle, 'touchend', x, y);
  // Tap 2 within the double-tap window, then pan past the 5px threshold
  await page.waitForTimeout(100);
  await dispatchTouch(page, handle, 'touchstart', x, y);
  await dispatchTouch(page, handle, 'touchmove', x + 30, y + 60);
  await dispatchTouch(page, handle, 'touchmove', x + 30, y + 70);
  await dispatchTouch(page, handle, 'touchend', x + 30, y + 70);

  const after = await getWinState(page);
  expect(after.maximized).toBe(false);
  expect(after.width).toBe(before.width);
  expect(after.height).toBe(before.height);
  expect(after.top).toBeGreaterThan(maxState.top);
});

test('single-tab dock pane plain drag on maximized dock restores size and moves it', async ({ page }) => {
  await openApp(page);
  await uploadFile(page, 'sample1.csv');
  await waitForWindow(page, 'sample1');
  await executeSQL(page, 'SELECT * INTO [t2] FROM [sample1]');
  await waitForWindow(page, 't2');

  // Split dock: two single-tab leaves; plain tab drag moves the whole dock
  const preDock = await page.evaluate(() => {
    const w = app._test.windows;
    app._test.mergeWindowsAsSplit(w[0].id, w[1].id, 'right');
    const dock = app._test.dockContainers[0];
    return { width: dock.el.offsetWidth, height: dock.el.offsetHeight };
  });

  await page.evaluate(() => app._test.toggleMaximizeDock(app._test.dockContainers[0]));

  const firstTab = page.locator('.dock-tab').first();
  const box = await firstTab.boundingBox();
  const grabX = box.x + box.width / 2;
  const grabY = box.y + box.height / 2;
  await page.mouse.move(grabX, grabY);
  await page.mouse.down();
  await page.mouse.move(grabX + 30, grabY + 70, { steps: 5 });
  await page.mouse.up();

  const after = await page.evaluate(() => {
    const dock = app._test.dockContainers[0];
    return {
      maximized: dock.maximized,
      top: parseInt(dock.el.style.top),
      width: dock.el.offsetWidth,
      height: dock.el.offsetHeight,
      leaves: document.querySelectorAll('.dock-leaf').length,
    };
  });
  expect(after.maximized).toBe(false);
  expect(after.width).toBe(preDock.width);
  expect(after.height).toBe(preDock.height);
  expect(after.top).toBeGreaterThan(0);
  // Drag moved the dock, it did not undock or reorder anything
  expect(after.leaves).toBe(2);
});

test('resizing a maximized window clears the flag but keeps position', async ({ page }) => {
  await openApp(page);
  await uploadFile(page, 'sample1.csv');
  await waitForWindow(page, 'sample1');

  await page.evaluate(() => app._test.toggleMaximize(app._test.windows[0].id));

  // Drag the bottom-right resize handle inward
  const handle = page.locator('.subwindow .resize-handle.rh-br').first();
  const box = await handle.boundingBox();
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  await page.mouse.down();
  await page.mouse.move(box.x - 100, box.y - 80, { steps: 5 });
  await page.mouse.up();

  const after = await getWinState(page);
  // Resize is not a move: window unmaximizes logically but stays anchored
  expect(after.maximized).toBe(false);
  expect(after.left).toBe(0);
  expect(after.top).toBe(0);
});
