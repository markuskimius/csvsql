const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, executeSQL } = require('../helpers');

// Regression tests for standalone/docked feature parity.
test.describe('Dock parity', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, 'sample1.csv');
    await waitForWindow(page, 'sample1');
    await executeSQL(page, "SELECT * INTO [query_result] FROM [sample1]");
    await waitForWindow(page, 'query_result');
  });

  test('DROP TABLE on a docked tab removes the tab and dissolves the dock', async ({ page }) => {
    await page.evaluate(() => {
      const wins = app._test.windows;
      app._test.mergeWindowsIntoTabs(wins[0].id, wins[1].id);
    });
    await page.waitForTimeout(100);

    await executeSQL(page, 'DROP TABLE [query_result]');
    await page.waitForTimeout(200);

    const result = await page.evaluate(() => ({
      dockCount: app._test.dockContainers.length,
      windowCount: app._test.windows.length,
      tabCount: document.querySelectorAll('.dock-tab').length,
      orphanBodies: document.querySelectorAll('.dock-leaf .win-body').length,
      hasQueryResult: !!app._test.tables['query_result'],
      survivorStandalone: app._test.windows.every(w => !w.dockId),
    }));
    // Removing one of two tabs dissolves the dock back to a standalone window
    expect(result.dockCount).toBe(0);
    expect(result.windowCount).toBe(1);
    expect(result.tabCount).toBe(0);
    expect(result.orphanBodies).toBe(0);
    expect(result.hasQueryResult).toBe(false);
    expect(result.survivorStandalone).toBe(true);

    // Survivor still works: run a query against it
    await executeSQL(page, 'SELECT * INTO [check1] FROM [sample1]');
    await waitForWindow(page, 'check1');
  });

  test('DROP TABLE on one tab of three keeps the dock with remaining tabs', async ({ page }) => {
    await executeSQL(page, "SELECT * INTO [third] FROM [sample1]");
    await waitForWindow(page, 'third');

    await page.evaluate(() => {
      const wins = app._test.windows;
      app._test.mergeWindowsIntoTabs(wins[0].id, wins[1].id);
      const dock = app._test.dockContainers[0];
      app._test.dockWindowAsTab(wins[2].id, dock.root, dock);
    });
    await page.waitForTimeout(100);

    await executeSQL(page, 'DROP TABLE [query_result]');
    await page.waitForTimeout(200);

    const result = await page.evaluate(() => ({
      dockCount: app._test.dockContainers.length,
      tabCount: document.querySelectorAll('.dock-tab').length,
      tabTitles: Array.from(document.querySelectorAll('.dock-tab-title')).map(t => t.textContent),
      bodiesInLeaf: document.querySelectorAll('.dock-leaf .win-body').length,
    }));
    expect(result.dockCount).toBe(1);
    expect(result.tabCount).toBe(2);
    expect(result.tabTitles.join(' ')).not.toContain('query_result');
    expect(result.bodiesInLeaf).toBe(2);
  });

  test('splitter drag closes an open autofilter dropdown', async ({ page }) => {
    await page.evaluate(() => {
      const wins = app._test.windows;
      app._test.mergeWindowsAsSplit(wins[0].id, wins[1].id, 'right');
    });
    await page.waitForTimeout(100);

    await page.click('.dock-leaf .col-filter-btn');
    await expect(page.locator('.autofilter-dropdown')).toBeVisible();

    const splitter = page.locator('.dock-splitter');
    const box = await splitter.boundingBox();
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + box.width / 2 + 30, box.y + box.height / 2);
    await page.mouse.up();

    await expect(page.locator('.autofilter-dropdown')).toHaveCount(0);
  });

  test('dock resize closes an open autofilter dropdown', async ({ page }) => {
    await page.evaluate(() => {
      const wins = app._test.windows;
      app._test.mergeWindowsIntoTabs(wins[0].id, wins[1].id);
    });
    await page.waitForTimeout(100);

    await page.locator('.dock-leaf .win-body:visible .col-filter-btn').first().click();
    await expect(page.locator('.autofilter-dropdown')).toBeVisible();

    const handle = page.locator('.dock-container .resize-handle.rh-br');
    const box = await handle.boundingBox();
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + 40, box.y + 40);
    await page.mouse.up();

    await expect(page.locator('.autofilter-dropdown')).toHaveCount(0);
  });

  test('Ctrl+Shift+L cycles into a docked inactive tab and activates it', async ({ page }) => {
    await executeSQL(page, "SELECT * INTO [third] FROM [sample1]");
    await waitForWindow(page, 'third');

    const setup = await page.evaluate(() => {
      const wins = app._test.windows;
      // Dock first two windows as tabs; third stays standalone and focused
      app._test.mergeWindowsIntoTabs(wins[0].id, wins[1].id);
      const standalone = wins.find(w => !w.dockId);
      app._test.focusWindow(standalone.id);
      return { standaloneId: standalone.id, dockedIds: wins.filter(w => w.dockId).map(w => w.id) };
    });

    // The shortcut fires from the per-table keydown handler, so focus a cell first
    await page.locator('.subwindow:not(.docked) .data-cell').first().click();
    await page.keyboard.press('Control+Shift+L');
    await page.waitForTimeout(100);
    const after1 = await page.evaluate(() => ({
      activeWinId: app._test.activeWinId,
      activeTab: app._test.dockContainers[0].root.activeTab,
    }));
    expect(setup.dockedIds).toContain(after1.activeWinId);
    expect(after1.activeTab).toBe(after1.activeWinId);

    // Cycle again from within the now-active docked tab: moves to the other tab
    await page.locator('.dock-leaf .win-body:visible .data-cell').first().click();
    await page.keyboard.press('Control+Shift+L');
    await page.waitForTimeout(100);
    const after2 = await page.evaluate(() => ({
      activeWinId: app._test.activeWinId,
      activeTab: app._test.dockContainers[0].root.activeTab,
    }));
    expect(setup.dockedIds).toContain(after2.activeWinId);
    expect(after2.activeWinId).not.toBe(after1.activeWinId);
    expect(after2.activeTab).toBe(after2.activeWinId);
  });

  test('Ctrl+L nudges the dock container when the active window is docked', async ({ page }) => {
    const before = await page.evaluate(() => {
      const wins = app._test.windows;
      app._test.mergeWindowsIntoTabs(wins[0].id, wins[1].id);
      const dock = app._test.dockContainers[0];
      app._test.focusWindow(wins[0].id);
      return { left: parseInt(dock.el.style.left), top: parseInt(dock.el.style.top) };
    });

    // The shortcut fires from the per-table keydown handler, so focus a cell first
    await page.locator('.dock-leaf .win-body:visible .data-cell').first().click();
    await page.keyboard.press('Control+l');
    await page.waitForTimeout(100);

    const after = await page.evaluate(() => {
      const dock = app._test.dockContainers[0];
      return { left: parseInt(dock.el.style.left), top: parseInt(dock.el.style.top) };
    });
    expect(after.left).toBe(before.left + 5);
    expect(after.top).toBe(before.top);
  });

  test('undocking a tab from a maximized dock yields a non-maximized window', async ({ page }) => {
    const result = await page.evaluate(() => {
      const wins = app._test.windows;
      app._test.mergeWindowsIntoTabs(wins[0].id, wins[1].id);
      const dock = app._test.dockContainers[0];
      app._test.toggleMaximizeDock(dock);
      const undockId = wins[0].id;
      app._test.undockWindow(undockId);
      const win = wins.find(w => w.id === undockId);
      return {
        maximized: !!win.maximized,
        prevBounds: win.prevBounds || null,
        dockId: win.dockId,
      };
    });
    expect(result.maximized).toBe(false);
    expect(result.prevBounds).toBeNull();
    expect(result.dockId).toBeNull();
  });
});
