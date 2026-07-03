const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, executeSQL } = require('../helpers');

test.describe('Tab Vertical Undock', () => {

  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, 'sample1.csv');
    await waitForWindow(page, 'sample1');
    await executeSQL(page, "SELECT * INTO [t2] FROM [sample1]");
    await waitForWindow(page, 't2');
    await executeSQL(page, "SELECT * INTO [t3] FROM [sample1]");
    await waitForWindow(page, 't3');

    // Merge all three windows into a tab group
    await page.evaluate(() => {
      const wins = app._test.windows;
      app._test.mergeWindowsIntoTabs(wins[0].id, wins[1].id);
      const dock = app._test.dockContainers[0];
      app._test.dockWindowAsTab(wins[2].id, dock.root, dock);
    });

    // Verify setup
    const setup = await page.evaluate(() => ({
      dockCount: app._test.dockContainers.length,
      tabCount: app._test.dockContainers[0].root.tabs.length,
    }));
    expect(setup.dockCount).toBe(1);
    expect(setup.tabCount).toBe(3);
  });

  // =========================================================================
  // 1. Vertical drag past threshold undocks a tab
  // =========================================================================

  test('vertical drag past threshold undocks tab', async ({ page }) => {
    const tab = page.locator('.dock-tab').first();
    const tabBox = await tab.boundingBox();

    await page.mouse.move(tabBox.x + tabBox.width / 2, tabBox.y + tabBox.height / 2);
    await page.mouse.down({ button: 'left' });
    // Move past 5px threshold to enter reorder mode
    await page.mouse.move(
      tabBox.x + tabBox.width / 2,
      tabBox.y + tabBox.height / 2 + 8,
      { steps: 3 }
    );
    // Move past 30px vertical threshold to trigger undock
    await page.mouse.move(
      tabBox.x + tabBox.width / 2,
      tabBox.y + tabBox.height / 2 + 35,
      { steps: 5 }
    );
    // Drop on empty space far from the dock
    const dockBox = await page.locator('.dock-container').boundingBox();
    await page.mouse.move(
      dockBox.x + dockBox.width + 100,
      dockBox.y + dockBox.height + 100,
      { steps: 5 }
    );
    await page.mouse.up({ button: 'left' });

    const result = await page.evaluate(() => {
      const undocked = app._test.windows.filter(w => w.dockId === null);
      return {
        undockedCount: undocked.length,
        dockTabCount: app._test.dockContainers[0]?.root.tabs.length || 0,
      };
    });

    expect(result.undockedCount).toBe(1);
    expect(result.dockTabCount).toBe(2);
  });

  // =========================================================================
  // 2. Vertical drag creates ghost element
  // =========================================================================

  test('vertical drag creates ghost element', async ({ page }) => {
    const tab = page.locator('.dock-tab').first();
    const tabBox = await tab.boundingBox();

    await page.mouse.move(tabBox.x + tabBox.width / 2, tabBox.y + tabBox.height / 2);
    await page.mouse.down({ button: 'left' });
    // Move past 5px threshold
    await page.mouse.move(
      tabBox.x + tabBox.width / 2,
      tabBox.y + tabBox.height / 2 + 8,
      { steps: 3 }
    );
    // Move past 30px vertical threshold
    await page.mouse.move(
      tabBox.x + tabBox.width / 2,
      tabBox.y + tabBox.height / 2 + 35,
      { steps: 5 }
    );

    const ghostExists = await page.evaluate(() =>
      document.querySelector('body > .subwindow[style*="position: fixed"]') !== null
    );
    expect(ghostExists).toBe(true);

    await page.mouse.up({ button: 'left' });
  });

  // =========================================================================
  // 3. Vertical drag below threshold stays in reorder
  // =========================================================================

  test('vertical drag below threshold stays in reorder', async ({ page }) => {
    const beforeTabs = await page.evaluate(() =>
      app._test.dockContainers[0].root.tabs.slice()
    );

    const tab = page.locator('.dock-tab').first();
    const tabBox = await tab.boundingBox();

    await page.mouse.move(tabBox.x + tabBox.width / 2, tabBox.y + tabBox.height / 2);
    await page.mouse.down({ button: 'left' });
    // Move past 5px threshold but stay under 30px vertical
    await page.mouse.move(
      tabBox.x + tabBox.width / 2 + 50,
      tabBox.y + tabBox.height / 2 + 25,
      { steps: 5 }
    );

    // No ghost should exist
    const ghostExists = await page.evaluate(() =>
      document.querySelector('body > .subwindow[style*="position: fixed"]') !== null
    );
    expect(ghostExists).toBe(false);

    await page.mouse.up({ button: 'left' });

    // All 3 tabs should still be in the dock
    const afterTabs = await page.evaluate(() =>
      app._test.dockContainers[0].root.tabs.slice()
    );
    expect(afterTabs.length).toBe(3);
    // Same tabs present (order may differ from reorder, but count is preserved)
    expect(afterTabs.sort()).toEqual(beforeTabs.sort());
  });

  // =========================================================================
  // 4. Horizontal drag stays in reorder mode
  // =========================================================================

  test('horizontal drag stays in reorder mode', async ({ page }) => {
    const tabs = page.locator('.dock-tab');
    const firstTab = tabs.nth(0);
    const lastTab = tabs.nth(2);
    const firstBox = await firstTab.boundingBox();
    const lastBox = await lastTab.boundingBox();

    await page.mouse.move(firstBox.x + firstBox.width / 2, firstBox.y + firstBox.height / 2);
    await page.mouse.down({ button: 'left' });
    // Move horizontally past threshold, same y
    await page.mouse.move(
      lastBox.x + lastBox.width / 2,
      firstBox.y + firstBox.height / 2,
      { steps: 5 }
    );

    // No ghost should appear
    const ghostExists = await page.evaluate(() =>
      document.querySelector('body > .subwindow[style*="position: fixed"]') !== null
    );
    expect(ghostExists).toBe(false);

    // dock-tab-dragging class should be present (reorder mode active)
    const hasDragging = await page.evaluate(() =>
      document.querySelector('.dock-tab-dragging') !== null
    );
    expect(hasDragging).toBe(true);

    await page.mouse.up({ button: 'left' });

    // All tabs still in dock
    const tabCount = await page.evaluate(() =>
      app._test.dockContainers[0].root.tabs.length
    );
    expect(tabCount).toBe(3);
  });

  // =========================================================================
  // 5. Shift+drag still works for immediate undock
  // =========================================================================

  test('shift+drag still works for immediate undock', async ({ page }) => {
    const tab = page.locator('.dock-tab').first();
    const tabBox = await tab.boundingBox();
    const dockBox = await page.locator('.dock-container').boundingBox();

    await page.mouse.move(tabBox.x + tabBox.width / 2, tabBox.y + tabBox.height / 2);
    await page.keyboard.down('Shift');
    await page.mouse.down({ button: 'left' });
    // Move only 10px -- well below the 30px vertical threshold, but Shift bypasses it
    await page.mouse.move(
      tabBox.x + tabBox.width / 2 + 10,
      tabBox.y + tabBox.height / 2,
      { steps: 3 }
    );

    // Ghost should appear immediately with Shift held
    const ghostExists = await page.evaluate(() =>
      document.querySelector('body > .subwindow[style*="position: fixed"]') !== null
    );
    expect(ghostExists).toBe(true);

    // Drop outside the dock
    await page.mouse.move(
      dockBox.x + dockBox.width + 100,
      dockBox.y + dockBox.height + 100,
      { steps: 5 }
    );
    await page.mouse.up({ button: 'left' });
    await page.keyboard.up('Shift');

    const result = await page.evaluate(() => {
      const undocked = app._test.windows.filter(w => w.dockId === null);
      return {
        undockedCount: undocked.length,
        dockTabCount: app._test.dockContainers[0]?.root.tabs.length || 0,
      };
    });

    expect(result.undockedCount).toBe(1);
    expect(result.dockTabCount).toBe(2);
  });

  // =========================================================================
  // 6. Transition from reorder to undock cleans up reorder indicators
  // =========================================================================

  test('transition from reorder to undock cleans up reorder indicators', async ({ page }) => {
    const tabs = page.locator('.dock-tab');
    const firstTab = tabs.nth(0);
    const secondTab = tabs.nth(1);
    const firstBox = await firstTab.boundingBox();
    const secondBox = await secondTab.boundingBox();

    await page.mouse.move(firstBox.x + firstBox.width / 2, firstBox.y + firstBox.height / 2);
    await page.mouse.down({ button: 'left' });
    // Move horizontally to enter reorder mode
    await page.mouse.move(
      secondBox.x + secondBox.width / 2,
      firstBox.y + firstBox.height / 2,
      { steps: 5 }
    );

    // Verify reorder mode is active
    const reorderActive = await page.evaluate(() =>
      document.querySelector('.dock-tab-dragging') !== null
    );
    expect(reorderActive).toBe(true);

    // Now move vertically past 30px threshold to trigger undock transition
    await page.mouse.move(
      secondBox.x + secondBox.width / 2,
      firstBox.y + firstBox.height / 2 + 35,
      { steps: 5 }
    );

    // Reorder indicators should be cleaned up
    const anyReorderIndicators = await page.evaluate(() =>
      document.querySelector('.dock-tab-dragging, .dock-tab-drop-left, .dock-tab-drop-right') !== null
    );
    expect(anyReorderIndicators).toBe(false);

    // Ghost should now exist
    const ghostExists = await page.evaluate(() =>
      document.querySelector('body > .subwindow[style*="position: fixed"]') !== null
    );
    expect(ghostExists).toBe(true);

    await page.mouse.up({ button: 'left' });
  });

  // =========================================================================
  // 7. Undocked tab can be dropped on empty space
  // =========================================================================

  test('undocked tab can be dropped on empty space', async ({ page }) => {
    const tab = page.locator('.dock-tab').first();
    const tabBox = await tab.boundingBox();

    const undockedWinId = await page.evaluate(() =>
      app._test.dockContainers[0].root.tabs[0]
    );

    const dockBox = await page.locator('.dock-container').boundingBox();
    const dropX = dockBox.x + dockBox.width + 80;
    const dropY = dockBox.y + 50;

    await page.mouse.move(tabBox.x + tabBox.width / 2, tabBox.y + tabBox.height / 2);
    await page.mouse.down({ button: 'left' });
    // Move past 5px threshold
    await page.mouse.move(
      tabBox.x + tabBox.width / 2,
      tabBox.y + tabBox.height / 2 + 8,
      { steps: 3 }
    );
    // Move past 30px vertical threshold
    await page.mouse.move(
      tabBox.x + tabBox.width / 2,
      tabBox.y + tabBox.height / 2 + 35,
      { steps: 5 }
    );
    // Move to empty workspace area
    await page.mouse.move(dropX, dropY, { steps: 5 });
    await page.mouse.up({ button: 'left' });

    const result = await page.evaluate((winId) => {
      const win = app._test.windows.find(w => w.id === winId);
      return {
        isDocked: win.dockId !== null,
        left: parseInt(win.el.style.left),
        top: parseInt(win.el.style.top),
      };
    }, undockedWinId);

    expect(result.isDocked).toBe(false);
    // Window should be placed near the drop position (within the workspace area)
    expect(result.left).toBeGreaterThan(0);
    expect(result.top).toBeGreaterThanOrEqual(0);
  });

  // =========================================================================
  // 8. Undocked tab can be docked into another window
  // =========================================================================

  test('undocked tab can be docked into another window via vertical drag', async ({ page }) => {
    // Create a 4th table that stays standalone
    await executeSQL(page, "SELECT * INTO [t4] FROM [sample1]");
    await waitForWindow(page, 't4');

    // Position the dock and standalone window so they don't overlap
    await page.evaluate(() => {
      const dock = app._test.dockContainers[0];
      dock.el.style.left = '0px';
      dock.el.style.top = '0px';
      dock.el.style.width = '400px';
      dock.el.style.height = '300px';
      const standaloneWin = app._test.windows.find(w => w.dockId === null);
      standaloneWin.el.style.left = '420px';
      standaloneWin.el.style.top = '0px';
      standaloneWin.el.style.width = '300px';
      standaloneWin.el.style.height = '300px';
    });

    const standaloneInfo = await page.evaluate(() => {
      const standaloneWin = app._test.windows.find(w => w.dockId === null);
      return {
        id: standaloneWin.id,
        rect: standaloneWin.el.getBoundingClientRect(),
      };
    });

    const tab = page.locator('.dock-tab').first();
    const tabBox = await tab.boundingBox();

    const tabWinId = await page.evaluate(() =>
      app._test.dockContainers[0].root.tabs[0]
    );

    // Get titlebar area of the standalone window (center zone = titlebar)
    const targetTitleY = standaloneInfo.rect.top + 16;
    const targetX = standaloneInfo.rect.left + standaloneInfo.rect.width / 2;

    await page.mouse.move(tabBox.x + tabBox.width / 2, tabBox.y + tabBox.height / 2);
    await page.mouse.down({ button: 'left' });
    // Move vertically past threshold
    await page.mouse.move(
      tabBox.x + tabBox.width / 2,
      tabBox.y + tabBox.height / 2 + 35,
      { steps: 5 }
    );
    // Move onto the standalone window's titlebar (center zone triggers tabbing)
    await page.mouse.move(targetX, targetTitleY, { steps: 5 });
    await page.mouse.up({ button: 'left' });

    const result = await page.evaluate((winId) => {
      const win = app._test.windows.find(w => w.id === winId);
      return {
        isDocked: win.dockId !== null,
      };
    }, tabWinId);

    // The tab should now be docked (merged with the standalone window into a new dock)
    expect(result.isDocked).toBe(true);
  });

  // =========================================================================
  // 9. Ghost removed on mouseup
  // =========================================================================

  test('ghost removed on mouseup', async ({ page }) => {
    const tab = page.locator('.dock-tab').first();
    const tabBox = await tab.boundingBox();

    await page.mouse.move(tabBox.x + tabBox.width / 2, tabBox.y + tabBox.height / 2);
    await page.mouse.down({ button: 'left' });
    // Move past 5px threshold
    await page.mouse.move(
      tabBox.x + tabBox.width / 2,
      tabBox.y + tabBox.height / 2 + 8,
      { steps: 3 }
    );
    // Move past 30px vertical threshold
    await page.mouse.move(
      tabBox.x + tabBox.width / 2,
      tabBox.y + tabBox.height / 2 + 35,
      { steps: 5 }
    );

    // Ghost should exist before mouseup
    const ghostBefore = await page.evaluate(() =>
      document.querySelector('body > .subwindow[style*="position: fixed"]') !== null
    );
    expect(ghostBefore).toBe(true);

    await page.mouse.up({ button: 'left' });

    // Ghost should be removed after mouseup
    const ghostAfter = await page.evaluate(() =>
      document.querySelector('body > .subwindow[style*="position: fixed"]') !== null
    );
    expect(ghostAfter).toBe(false);
  });
});
