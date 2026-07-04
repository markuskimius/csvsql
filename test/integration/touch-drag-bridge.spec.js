const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, executeSQL, dispatchTouch, touchCenterOf } = require('../helpers');

// The touch drag bridge re-dispatches touch sequences on drag handles as
// synthetic mouse events, so every mousedown-driven drag works by touch.

async function touchDrag(page, locator, steps) {
  const handle = await locator.elementHandle();
  const { x, y } = await touchCenterOf(locator);
  await dispatchTouch(page, handle, 'touchstart', x, y);
  for (const [dx, dy] of steps) {
    await dispatchTouch(page, handle, 'touchmove', x + dx, y + dy);
  }
  const [lx, ly] = steps.length ? steps[steps.length - 1] : [0, 0];
  await dispatchTouch(page, handle, 'touchend', x + lx, y + ly);
}

test.describe('Touch drag bridge', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, 'sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('touch drag on a corner handle resizes the window', async ({ page }) => {
    const win = page.locator('.subwindow').first();
    const before = await win.boundingBox();
    await touchDrag(page, win.locator('.resize-handle.rh-br'), [[20, 15], [60, 40], [80, 60]]);
    const after = await win.boundingBox();
    expect(after.width).toBeGreaterThan(before.width + 50);
    expect(after.height).toBeGreaterThan(before.height + 30);
  });

  test('touch drag on a column resize handle widens the column', async ({ page }) => {
    const before = await page.evaluate(() => [...app._test.windows[0].colWidths]);
    const handle = page.locator('.subwindow th[data-col-idx="0"] .col-resize-handle:not(.col-resize-left)');
    await touchDrag(page, handle, [[10, 0], [30, 0], [50, 0]]);
    const after = await page.evaluate(() => [...app._test.windows[0].colWidths]);
    expect(after[0]).toBeGreaterThan(before[0] + 30);
    expect(after[1]).toBe(before[1]);
  });

  test('touch drag resizes the SQL console panel', async ({ page }) => {
    const panel = page.locator('#console-panel');
    const before = await panel.boundingBox();
    await touchDrag(page, page.locator('#console-resize-handle'), [[0, -20], [0, -40], [0, -60]]);
    const after = await panel.boundingBox();
    expect(after.height).toBeGreaterThan(before.height + 40);
  });

  test('touch pan on row numbers drag-selects rows', async ({ page }) => {
    const firstRowNum = page.locator('.subwindow td.row-num').first();
    const handle = await firstRowNum.elementHandle();
    const { x, y } = await touchCenterOf(firstRowNum);
    await dispatchTouch(page, handle, 'touchstart', x, y);
    await dispatchTouch(page, handle, 'touchmove', x, y + 26);
    await dispatchTouch(page, handle, 'touchmove', x, y + 52);
    await dispatchTouch(page, handle, 'touchend', x, y + 52);
    const result = await page.evaluate(() => {
      const win = app._test.windows[0];
      return { selected: win.selectedCells.size, cols: app._test.tables['sample1'].columns.length };
    });
    expect(result.selected).toBe(3 * result.cols);
  });

  test('touch tap on a row number selects that row', async ({ page }) => {
    const firstRowNum = page.locator('.subwindow td.row-num').first();
    const handle = await firstRowNum.elementHandle();
    const { x, y } = await touchCenterOf(firstRowNum);
    await dispatchTouch(page, handle, 'touchstart', x, y);
    await dispatchTouch(page, handle, 'touchend', x, y);
    const result = await page.evaluate(() => {
      const win = app._test.windows[0];
      return { selected: win.selectedCells.size, cols: app._test.tables['sample1'].columns.length };
    });
    expect(result.selected).toBe(result.cols);
  });

  test.describe('with a dock', () => {
    test.beforeEach(async ({ page }) => {
      await executeSQL(page, 'SELECT * INTO [query_result] FROM [sample1]');
      await waitForWindow(page, 'query_result');
    });

    test('vertical touch drag on a tab undocks it', async ({ page }) => {
      await page.evaluate(() => {
        const wins = app._test.windows;
        app._test.mergeWindowsIntoTabs(wins[0].id, wins[1].id);
      });
      await page.waitForTimeout(100);
      const tab = page.locator('.dock-tab').first();
      const handle = await tab.elementHandle();
      const { x, y } = await touchCenterOf(tab);
      const area = await page.locator('#window-area').boundingBox();
      const dropX = area.x + area.width - 80;
      const dropY = area.y + area.height - 80;
      await dispatchTouch(page, handle, 'touchstart', x, y);
      // Tab DOM can be re-rendered mid-drag; the bridge listens on document,
      // so dispatch subsequent events on a stable element.
      const body = await page.evaluateHandle(() => document.body);
      await dispatchTouch(page, body, 'touchmove', x, y + 10);
      await dispatchTouch(page, body, 'touchmove', x, y + 60);
      await dispatchTouch(page, body, 'touchmove', dropX, dropY);
      await dispatchTouch(page, body, 'touchend', dropX, dropY);
      await page.waitForTimeout(100);
      const result = await page.evaluate(() => ({
        dockCount: app._test.dockContainers.length,
        standalone: app._test.windows.filter(w => !w.dockId).length,
      }));
      expect(result.dockCount).toBe(0);
      expect(result.standalone).toBe(2);
    });

    test('touch drag on the splitter changes the split ratio', async ({ page }) => {
      await page.evaluate(() => {
        const wins = app._test.windows;
        app._test.mergeWindowsAsSplit(wins[0].id, wins[1].id, 'right');
      });
      await page.waitForTimeout(100);
      const before = await page.evaluate(() => app._test.dockContainers[0].root.ratio);
      await touchDrag(page, page.locator('.dock-splitter'), [[20, 0], [50, 0], [80, 0]]);
      const after = await page.evaluate(() => app._test.dockContainers[0].root.ratio);
      expect(after).toBeGreaterThan(before + 0.05);
    });

    test('double-tap on the splitter resets the ratio to 0.5', async ({ page }) => {
      await page.evaluate(() => {
        const wins = app._test.windows;
        app._test.mergeWindowsAsSplit(wins[0].id, wins[1].id, 'right');
      });
      await page.waitForTimeout(100);
      const splitter = page.locator('.dock-splitter');
      await touchDrag(page, splitter, [[20, 0], [50, 0], [80, 0]]);
      const dragged = await page.evaluate(() => app._test.dockContainers[0].root.ratio);
      expect(dragged).not.toBeCloseTo(0.5, 2);
      // Two quick stationary taps → synthetic dblclick → ratio reset
      const handle = await splitter.elementHandle();
      const { x, y } = await touchCenterOf(splitter);
      await dispatchTouch(page, handle, 'touchstart', x, y);
      await dispatchTouch(page, handle, 'touchend', x, y);
      await page.waitForTimeout(80);
      await dispatchTouch(page, handle, 'touchstart', x, y);
      await dispatchTouch(page, handle, 'touchend', x, y);
      const after = await page.evaluate(() => app._test.dockContainers[0].root.ratio);
      expect(after).toBeCloseTo(0.5, 5);
    });

    test('touch tap on an inactive tab activates it', async ({ page }) => {
      await page.evaluate(() => {
        const wins = app._test.windows;
        app._test.mergeWindowsIntoTabs(wins[0].id, wins[1].id);
      });
      await page.waitForTimeout(100);
      const inactive = page.locator('.dock-tab:not(.active)').first();
      const winId = await inactive.getAttribute('data-win-id');
      const handle = await inactive.elementHandle();
      const { x, y } = await touchCenterOf(inactive);
      await dispatchTouch(page, handle, 'touchstart', x, y);
      await dispatchTouch(page, handle, 'touchend', x, y);
      await page.waitForTimeout(100);
      const activeId = await page.evaluate(() => {
        const t = document.querySelector('.dock-tab.active');
        return t ? t.dataset.winId : null;
      });
      expect(activeId).toBe(winId);
    });
  });
});
