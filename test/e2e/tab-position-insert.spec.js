const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, executeSQL } = require('../helpers');

test.describe('Tab Position Insert', () => {

  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, 'sample1.csv');
    await waitForWindow(page, 'sample1');
    await executeSQL(page, "SELECT * INTO [t2] FROM [sample1]");
    await waitForWindow(page, 't2');
    await executeSQL(page, "SELECT * INTO [t3] FROM [sample1]");
    await waitForWindow(page, 't3');
  });

  // =========================================================================
  // 1. dockWindowAsTab with atIndex inserts at position 0
  // =========================================================================
  test.describe('API-level dockWindowAsTab with atIndex', () => {

    test('atIndex 0 inserts at first position', async ({ page }) => {
      // Create a 2-tab dock from the first two windows
      await page.evaluate(() => {
        const wins = app._test.windows;
        app._test.mergeWindowsIntoTabs(wins[0].id, wins[1].id);
      });

      const result = await page.evaluate(() => {
        const wins = app._test.windows;
        const dock = app._test.dockContainers[0];
        const originalTabs = dock.root.tabs.slice();
        const thirdWin = wins[2];
        app._test.dockWindowAsTab(thirdWin.id, dock.root, dock, 0);
        return {
          tabs: dock.root.tabs.slice(),
          originalTabs,
          thirdWinId: thirdWin.id,
        };
      });

      // The third window should be first in the tab order
      expect(result.tabs[0]).toBe(result.thirdWinId);
      expect(result.tabs[1]).toBe(result.originalTabs[0]);
      expect(result.tabs[2]).toBe(result.originalTabs[1]);
    });

    // =========================================================================
    // 2. atIndex=1 inserts in the middle
    // =========================================================================

    test('atIndex 1 inserts in the middle', async ({ page }) => {
      await page.evaluate(() => {
        const wins = app._test.windows;
        app._test.mergeWindowsIntoTabs(wins[0].id, wins[1].id);
      });

      const result = await page.evaluate(() => {
        const wins = app._test.windows;
        const dock = app._test.dockContainers[0];
        const originalTabs = dock.root.tabs.slice();
        const thirdWin = wins[2];
        app._test.dockWindowAsTab(thirdWin.id, dock.root, dock, 1);
        return {
          tabs: dock.root.tabs.slice(),
          originalTabs,
          thirdWinId: thirdWin.id,
        };
      });

      expect(result.tabs[0]).toBe(result.originalTabs[0]);
      expect(result.tabs[1]).toBe(result.thirdWinId);
      expect(result.tabs[2]).toBe(result.originalTabs[1]);
    });

    // =========================================================================
    // 3. Without atIndex, appends to end
    // =========================================================================

    test('without atIndex appends to end', async ({ page }) => {
      await page.evaluate(() => {
        const wins = app._test.windows;
        app._test.mergeWindowsIntoTabs(wins[0].id, wins[1].id);
      });

      const result = await page.evaluate(() => {
        const wins = app._test.windows;
        const dock = app._test.dockContainers[0];
        const originalTabs = dock.root.tabs.slice();
        const thirdWin = wins[2];
        app._test.dockWindowAsTab(thirdWin.id, dock.root, dock);
        return {
          tabs: dock.root.tabs.slice(),
          originalTabs,
          thirdWinId: thirdWin.id,
        };
      });

      expect(result.tabs[0]).toBe(result.originalTabs[0]);
      expect(result.tabs[1]).toBe(result.originalTabs[1]);
      expect(result.tabs[2]).toBe(result.thirdWinId);
    });

    // =========================================================================
    // 4. atIndex beyond length appends to end
    // =========================================================================

    test('atIndex beyond length appends to end', async ({ page }) => {
      await page.evaluate(() => {
        const wins = app._test.windows;
        app._test.mergeWindowsIntoTabs(wins[0].id, wins[1].id);
      });

      const result = await page.evaluate(() => {
        const wins = app._test.windows;
        const dock = app._test.dockContainers[0];
        const originalTabs = dock.root.tabs.slice();
        const thirdWin = wins[2];
        app._test.dockWindowAsTab(thirdWin.id, dock.root, dock, 99);
        return {
          tabs: dock.root.tabs.slice(),
          originalTabs,
          thirdWinId: thirdWin.id,
        };
      });

      expect(result.tabs[0]).toBe(result.originalTabs[0]);
      expect(result.tabs[1]).toBe(result.originalTabs[1]);
      expect(result.tabs[2]).toBe(result.thirdWinId);
    });
  });

  // =========================================================================
  // 5-10. Mouse-driven position-aware insertion
  // =========================================================================
  test.describe('Mouse-driven tab insertion', () => {

    test.beforeEach(async ({ page }) => {
      // Create a 3-tab dock from the three windows
      await page.evaluate(() => {
        const wins = app._test.windows;
        app._test.mergeWindowsIntoTabs(wins[0].id, wins[1].id);
        const dock = app._test.dockContainers[0];
        app._test.dockWindowAsTab(wins[2].id, dock.root, dock);
      });

      const setup = await page.evaluate(() => ({
        dockCount: app._test.dockContainers.length,
        tabCount: app._test.dockContainers[0].root.tabs.length,
      }));
      expect(setup.dockCount).toBe(1);
      expect(setup.tabCount).toBe(3);
    });

    // =========================================================================
    // 5. Drop indicator appears during ghost drag over tab bar
    // =========================================================================

    test('drop indicator appears during ghost drag over tab bar', async ({ page }) => {
      // Create a 4th standalone window
      await executeSQL(page, "SELECT * INTO [t4] FROM [sample1]");
      await waitForWindow(page, 't4');

      // Get standalone window's titlebar for shift+drag
      const standaloneWin = await page.evaluate(() => {
        const win = app._test.windows.find(w => w.dockId === null);
        const rect = win.el.getBoundingClientRect();
        return { id: win.id, left: rect.left, top: rect.top, width: rect.width, height: rect.height };
      });

      // Get the dock's tab bar area
      const tabBar = page.locator('.dock-tab-bar').first();
      const tabBarBox = await tabBar.boundingBox();

      // Shift+drag the standalone window over the dock tab bar
      const startX = standaloneWin.left + standaloneWin.width / 2;
      const startY = standaloneWin.top + 16; // titlebar center
      await page.mouse.move(startX, startY);
      await page.keyboard.down('Shift');
      await page.mouse.down({ button: 'left' });
      // Move past threshold
      await page.mouse.move(startX + 10, startY, { steps: 3 });
      // Move onto the dock tab bar (center zone)
      await page.mouse.move(
        tabBarBox.x + tabBarBox.width / 2,
        tabBarBox.y + tabBarBox.height / 2,
        { steps: 5 }
      );

      const indicatorExists = await page.evaluate(() =>
        document.querySelector('.dock-drop-tab-indicator') !== null
      );
      expect(indicatorExists).toBe(true);

      await page.mouse.up({ button: 'left' });
      await page.keyboard.up('Shift');
    });

    // =========================================================================
    // 6. Drop indicator position updates as cursor moves
    // =========================================================================

    test('drop indicator position updates as cursor moves', async ({ page }) => {
      await executeSQL(page, "SELECT * INTO [t4] FROM [sample1]");
      await waitForWindow(page, 't4');

      const standaloneWin = await page.evaluate(() => {
        const win = app._test.windows.find(w => w.dockId === null);
        const rect = win.el.getBoundingClientRect();
        return { left: rect.left, top: rect.top, width: rect.width, height: rect.height };
      });

      const tabBar = page.locator('.dock-tab-bar').first();
      const tabBarBox = await tabBar.boundingBox();
      const firstTab = page.locator('.dock-tab').nth(0);
      const lastTab = page.locator('.dock-tab').nth(2);
      const firstTabBox = await firstTab.boundingBox();
      const lastTabBox = await lastTab.boundingBox();

      // Shift+drag from standalone window
      const startX = standaloneWin.left + standaloneWin.width / 2;
      const startY = standaloneWin.top + 16;
      await page.mouse.move(startX, startY);
      await page.keyboard.down('Shift');
      await page.mouse.down({ button: 'left' });
      await page.mouse.move(startX + 10, startY, { steps: 3 });

      // Move to left side of tab bar (before first tab)
      await page.mouse.move(
        firstTabBox.x + 2,
        tabBarBox.y + tabBarBox.height / 2,
        { steps: 5 }
      );

      const leftIndicator = await page.evaluate(() => {
        const ind = document.querySelector('.dock-drop-tab-indicator');
        return ind ? parseFloat(ind.style.left) : null;
      });

      // Move to right side (after last tab)
      await page.mouse.move(
        lastTabBox.x + lastTabBox.width + 5,
        tabBarBox.y + tabBarBox.height / 2,
        { steps: 5 }
      );

      const rightIndicator = await page.evaluate(() => {
        const ind = document.querySelector('.dock-drop-tab-indicator');
        return ind ? parseFloat(ind.style.left) : null;
      });

      // The indicator's left position should have changed
      expect(leftIndicator).not.toBeNull();
      expect(rightIndicator).not.toBeNull();
      expect(rightIndicator).toBeGreaterThan(leftIndicator);

      await page.mouse.up({ button: 'left' });
      await page.keyboard.up('Shift');
    });

    // =========================================================================
    // 7. Tab inserted at correct position (index 0) via mouse drag
    // =========================================================================

    test('tab inserted at index 0 via mouse drag', async ({ page }) => {
      await executeSQL(page, "SELECT * INTO [t4] FROM [sample1]");
      await waitForWindow(page, 't4');

      const standaloneWinId = await page.evaluate(() =>
        app._test.windows.find(w => w.dockId === null).id
      );

      const standaloneWin = await page.evaluate(() => {
        const win = app._test.windows.find(w => w.dockId === null);
        const rect = win.el.getBoundingClientRect();
        return { left: rect.left, top: rect.top, width: rect.width, height: rect.height };
      });

      const firstTab = page.locator('.dock-tab').nth(0);
      const firstTabBox = await firstTab.boundingBox();
      const tabBar = page.locator('.dock-tab-bar').first();
      const tabBarBox = await tabBar.boundingBox();

      const beforeTabs = await page.evaluate(() =>
        app._test.dockContainers[0].root.tabs.slice()
      );

      // Shift+drag the standalone window to the left of the first tab
      const startX = standaloneWin.left + standaloneWin.width / 2;
      const startY = standaloneWin.top + 16;
      await page.mouse.move(startX, startY);
      await page.keyboard.down('Shift');
      await page.mouse.down({ button: 'left' });
      await page.mouse.move(startX + 10, startY, { steps: 3 });
      // Drop on left side of first tab (before the midpoint)
      await page.mouse.move(
        firstTabBox.x + 2,
        tabBarBox.y + tabBarBox.height / 2,
        { steps: 5 }
      );
      await page.mouse.up({ button: 'left' });
      await page.keyboard.up('Shift');

      const afterTabs = await page.evaluate(() =>
        app._test.dockContainers[0].root.tabs.slice()
      );

      // The standalone window should be at index 0
      expect(afterTabs[0]).toBe(standaloneWinId);
      // Original tabs should follow
      expect(afterTabs[1]).toBe(beforeTabs[0]);
      expect(afterTabs[2]).toBe(beforeTabs[1]);
      expect(afterTabs[3]).toBe(beforeTabs[2]);
    });

    // =========================================================================
    // 8. Tab inserted at end when cursor past last tab
    // =========================================================================

    test('tab inserted at end when cursor past last tab', async ({ page }) => {
      await executeSQL(page, "SELECT * INTO [t4] FROM [sample1]");
      await waitForWindow(page, 't4');

      const standaloneWinId = await page.evaluate(() =>
        app._test.windows.find(w => w.dockId === null).id
      );

      const standaloneWin = await page.evaluate(() => {
        const win = app._test.windows.find(w => w.dockId === null);
        const rect = win.el.getBoundingClientRect();
        return { left: rect.left, top: rect.top, width: rect.width, height: rect.height };
      });

      const lastTab = page.locator('.dock-tab').nth(2);
      const lastTabBox = await lastTab.boundingBox();
      const tabBar = page.locator('.dock-tab-bar').first();
      const tabBarBox = await tabBar.boundingBox();

      const beforeTabs = await page.evaluate(() =>
        app._test.dockContainers[0].root.tabs.slice()
      );

      // Shift+drag to after the last tab
      const startX = standaloneWin.left + standaloneWin.width / 2;
      const startY = standaloneWin.top + 16;
      await page.mouse.move(startX, startY);
      await page.keyboard.down('Shift');
      await page.mouse.down({ button: 'left' });
      await page.mouse.move(startX + 10, startY, { steps: 3 });
      // Drop past the right edge of the last tab
      await page.mouse.move(
        lastTabBox.x + lastTabBox.width + 15,
        tabBarBox.y + tabBarBox.height / 2,
        { steps: 5 }
      );
      await page.mouse.up({ button: 'left' });
      await page.keyboard.up('Shift');

      const afterTabs = await page.evaluate(() =>
        app._test.dockContainers[0].root.tabs.slice()
      );

      // Original tabs should be at their original positions
      expect(afterTabs[0]).toBe(beforeTabs[0]);
      expect(afterTabs[1]).toBe(beforeTabs[1]);
      expect(afterTabs[2]).toBe(beforeTabs[2]);
      // The standalone window should be last
      expect(afterTabs[3]).toBe(standaloneWinId);
    });

    // =========================================================================
    // 9. Tab inserted in middle position
    // =========================================================================

    test('tab inserted in middle position', async ({ page }) => {
      await executeSQL(page, "SELECT * INTO [t4] FROM [sample1]");
      await waitForWindow(page, 't4');

      const standaloneWinId = await page.evaluate(() =>
        app._test.windows.find(w => w.dockId === null).id
      );

      const standaloneWin = await page.evaluate(() => {
        const win = app._test.windows.find(w => w.dockId === null);
        const rect = win.el.getBoundingClientRect();
        return { left: rect.left, top: rect.top, width: rect.width, height: rect.height };
      });

      const secondTab = page.locator('.dock-tab').nth(1);
      const secondTabBox = await secondTab.boundingBox();
      const tabBar = page.locator('.dock-tab-bar').first();
      const tabBarBox = await tabBar.boundingBox();

      const beforeTabs = await page.evaluate(() =>
        app._test.dockContainers[0].root.tabs.slice()
      );

      // Shift+drag to just before the second tab (left side of its midpoint)
      const startX = standaloneWin.left + standaloneWin.width / 2;
      const startY = standaloneWin.top + 16;
      await page.mouse.move(startX, startY);
      await page.keyboard.down('Shift');
      await page.mouse.down({ button: 'left' });
      await page.mouse.move(startX + 10, startY, { steps: 3 });
      // Position cursor on the left side of the second tab to insert at index 1
      await page.mouse.move(
        secondTabBox.x + 2,
        tabBarBox.y + tabBarBox.height / 2,
        { steps: 5 }
      );
      await page.mouse.up({ button: 'left' });
      await page.keyboard.up('Shift');

      const afterTabs = await page.evaluate(() =>
        app._test.dockContainers[0].root.tabs.slice()
      );

      // Original first tab stays first
      expect(afterTabs[0]).toBe(beforeTabs[0]);
      // New tab inserted at index 1
      expect(afterTabs[1]).toBe(standaloneWinId);
      // Original second and third tabs shift right
      expect(afterTabs[2]).toBe(beforeTabs[1]);
      expect(afterTabs[3]).toBe(beforeTabs[2]);
    });

    // =========================================================================
    // 10. Indicator cleaned up after drop
    // =========================================================================

    test('indicator cleaned up after drop', async ({ page }) => {
      await executeSQL(page, "SELECT * INTO [t4] FROM [sample1]");
      await waitForWindow(page, 't4');

      const standaloneWin = await page.evaluate(() => {
        const win = app._test.windows.find(w => w.dockId === null);
        const rect = win.el.getBoundingClientRect();
        return { left: rect.left, top: rect.top, width: rect.width, height: rect.height };
      });

      const tabBar = page.locator('.dock-tab-bar').first();
      const tabBarBox = await tabBar.boundingBox();

      // Shift+drag the standalone window over tab bar
      const startX = standaloneWin.left + standaloneWin.width / 2;
      const startY = standaloneWin.top + 16;
      await page.mouse.move(startX, startY);
      await page.keyboard.down('Shift');
      await page.mouse.down({ button: 'left' });
      await page.mouse.move(startX + 10, startY, { steps: 3 });
      await page.mouse.move(
        tabBarBox.x + tabBarBox.width / 2,
        tabBarBox.y + tabBarBox.height / 2,
        { steps: 5 }
      );

      // Indicator should be visible before drop
      const indicatorBeforeDrop = await page.evaluate(() =>
        document.querySelector('.dock-drop-tab-indicator') !== null
      );
      expect(indicatorBeforeDrop).toBe(true);

      // Drop
      await page.mouse.up({ button: 'left' });
      await page.keyboard.up('Shift');

      // Indicator should be cleaned up
      const indicatorAfterDrop = await page.evaluate(() =>
        document.querySelector('.dock-drop-tab-indicator') !== null
      );
      expect(indicatorAfterDrop).toBe(false);

      // Also verify the full overlay is gone
      const overlayAfterDrop = await page.evaluate(() =>
        document.querySelector('.dock-drop-overlay') !== null
      );
      expect(overlayAfterDrop).toBe(false);
    });
  });
});
