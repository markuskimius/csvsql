const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, executeSQL, dispatchTouch, touchCenterOf } = require('../helpers');

// Long-press (~600ms hold, no prior tap) on a standalone window's titlebar
// enters ghost/dock mode — the touch equivalent of Shift+drag.

test.describe('Touch docking via titlebar long-press', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, 'sample1.csv');
    await waitForWindow(page, 'sample1');
    await executeSQL(page, 'SELECT * INTO [query_result] FROM [sample1]');
    await waitForWindow(page, 'query_result');
  });

  test('hold titlebar, drag onto another window body, drop → split dock', async ({ page }) => {
    const source = page.locator('.subwindow', { hasNot: page.locator('.docked') }).nth(1);
    const titlebar = source.locator('.win-titlebar');
    const handle = await titlebar.elementHandle();
    const { x, y } = await touchCenterOf(titlebar);

    const targetBody = page.locator('.subwindow').first().locator('.win-body');
    const tb = await targetBody.boundingBox();
    // Left drop zone of the target body → horizontal split
    const dropX = tb.x + tb.width * 0.15;
    const dropY = tb.y + tb.height * 0.5;

    await dispatchTouch(page, handle, 'touchstart', x, y);
    await page.waitForTimeout(750);  // hold — dock mode arms at ~600ms
    await dispatchTouch(page, handle, 'touchmove', x + 10, y + 10);
    await dispatchTouch(page, handle, 'touchmove', (x + dropX) / 2, (y + dropY) / 2);
    await dispatchTouch(page, handle, 'touchmove', dropX, dropY);
    await dispatchTouch(page, handle, 'touchend', dropX, dropY);
    await page.waitForTimeout(150);

    const result = await page.evaluate(() => ({
      dockCount: app._test.dockContainers.length,
      docked: app._test.windows.filter(w => w.dockId !== null).length,
      splitExists: document.querySelector('.dock-split') !== null,
    }));
    expect(result.dockCount).toBe(1);
    expect(result.docked).toBe(2);
    expect(result.splitExists).toBe(true);
  });

  test('hold titlebar, drop on the other titlebar → tab group', async ({ page }) => {
    const source = page.locator('.subwindow').nth(1);
    const titlebar = source.locator('.win-titlebar');
    const handle = await titlebar.elementHandle();
    const { x, y } = await touchCenterOf(titlebar);

    const targetBar = page.locator('.subwindow').first().locator('.win-titlebar');
    const tbb = await targetBar.boundingBox();
    const dropX = tbb.x + tbb.width * 0.4;
    const dropY = tbb.y + tbb.height * 0.5;

    await dispatchTouch(page, handle, 'touchstart', x, y);
    await page.waitForTimeout(750);
    await dispatchTouch(page, handle, 'touchmove', x + 10, y + 10);
    await dispatchTouch(page, handle, 'touchmove', dropX, dropY);
    await dispatchTouch(page, handle, 'touchend', dropX, dropY);
    await page.waitForTimeout(150);

    const result = await page.evaluate(() => ({
      dockCount: app._test.dockContainers.length,
      tabCount: document.querySelectorAll('.dock-tab').length,
    }));
    expect(result.dockCount).toBe(1);
    expect(result.tabCount).toBe(2);
  });

  test('hold without moving, then release does not dock or move', async ({ page }) => {
    const source = page.locator('.subwindow').nth(1);
    const before = await source.boundingBox();
    const titlebar = source.locator('.win-titlebar');
    const handle = await titlebar.elementHandle();
    const { x, y } = await touchCenterOf(titlebar);

    await dispatchTouch(page, handle, 'touchstart', x, y);
    await page.waitForTimeout(750);
    await dispatchTouch(page, handle, 'touchend', x, y);
    await page.waitForTimeout(100);

    const after = await source.boundingBox();
    const dockCount = await page.evaluate(() => app._test.dockContainers.length);
    expect(dockCount).toBe(0);
    expect(after.x).toBeCloseTo(before.x, 0);
    expect(after.y).toBeCloseTo(before.y, 0);
  });

  test('1.5-tap move still works (short hold is not a long-press)', async ({ page }) => {
    const source = page.locator('.subwindow').nth(1);
    const before = await source.boundingBox();
    const titlebar = source.locator('.win-titlebar');
    const handle = await titlebar.elementHandle();
    const { x, y } = await touchCenterOf(titlebar);

    // Tap 1
    await dispatchTouch(page, handle, 'touchstart', x, y);
    await dispatchTouch(page, handle, 'touchend', x, y);
    await page.waitForTimeout(100);
    // Tap 2 + pan = move
    await dispatchTouch(page, handle, 'touchstart', x, y);
    await dispatchTouch(page, handle, 'touchmove', x + 40, y + 30);
    await dispatchTouch(page, handle, 'touchmove', x + 80, y + 60);
    await dispatchTouch(page, handle, 'touchend', x + 80, y + 60);
    await page.waitForTimeout(100);

    const after = await source.boundingBox();
    const dockCount = await page.evaluate(() => app._test.dockContainers.length);
    expect(dockCount).toBe(0);
    expect(after.x).toBeGreaterThan(before.x + 50);
  });
});
