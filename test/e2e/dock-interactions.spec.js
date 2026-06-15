const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, executeSQL } = require('../helpers');

test.describe('Dock Interactions', () => {

  // =========================================================================
  // 1. Left-click only: drag/resize operations require button 0
  // =========================================================================
  test.describe('Left-click only', () => {

    test.beforeEach(async ({ page }) => {
      await openApp(page);
      await uploadFile(page, 'sample1.csv');
      await waitForWindow(page, 'sample1');
    });

    test('right-click does NOT trigger window drag', async ({ page }) => {
      const origPos = await page.evaluate(() => {
        const win = app._test.windows[0];
        return { left: win.el.style.left, top: win.el.style.top };
      });

      const titlebar = page.locator('.subwindow .win-titlebar').first();
      const box = await titlebar.boundingBox();

      await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
      await page.mouse.down({ button: 'right' });
      await page.mouse.move(box.x + box.width / 2 + 50, box.y + box.height / 2 + 50);
      await page.mouse.up({ button: 'right' });

      const newPos = await page.evaluate(() => {
        const win = app._test.windows[0];
        return { left: win.el.style.left, top: win.el.style.top };
      });

      expect(newPos.left).toBe(origPos.left);
      expect(newPos.top).toBe(origPos.top);
    });

    test('right-click does NOT trigger window resize', async ({ page }) => {
      const origSize = await page.evaluate(() => {
        const win = app._test.windows[0];
        return { width: win.el.style.width, height: win.el.style.height };
      });

      const handle = page.locator('.subwindow .rh-br').first();
      const box = await handle.boundingBox();

      await page.mouse.move(box.x + 2, box.y + 2);
      await page.mouse.down({ button: 'right' });
      await page.mouse.move(box.x + 60, box.y + 60);
      await page.mouse.up({ button: 'right' });

      const newSize = await page.evaluate(() => {
        const win = app._test.windows[0];
        return { width: win.el.style.width, height: win.el.style.height };
      });

      expect(newSize.width).toBe(origSize.width);
      expect(newSize.height).toBe(origSize.height);
    });

    test('right-click does NOT trigger dock container resize', async ({ page }) => {
      await executeSQL(page, "SELECT * INTO [t2] FROM [sample1]");
      await waitForWindow(page, 't2');

      await page.evaluate(() => {
        const wins = app._test.windows;
        app._test.mergeWindowsIntoTabs(wins[0].id, wins[1].id);
      });

      const origSize = await page.evaluate(() => {
        const dock = app._test.dockContainers[0];
        return { width: dock.el.style.width, height: dock.el.style.height };
      });

      const handle = page.locator('.dock-container .rh-br').first();
      const box = await handle.boundingBox();

      await page.mouse.move(box.x + 2, box.y + 2);
      await page.mouse.down({ button: 'right' });
      await page.mouse.move(box.x + 60, box.y + 60);
      await page.mouse.up({ button: 'right' });

      const newSize = await page.evaluate(() => {
        const dock = app._test.dockContainers[0];
        return { width: dock.el.style.width, height: dock.el.style.height };
      });

      expect(newSize.width).toBe(origSize.width);
      expect(newSize.height).toBe(origSize.height);
    });

    test('right-click does NOT trigger splitter drag', async ({ page }) => {
      await executeSQL(page, "SELECT * INTO [t2] FROM [sample1]");
      await waitForWindow(page, 't2');

      await page.evaluate(() => {
        const wins = app._test.windows;
        app._test.mergeWindowsAsSplit(wins[0].id, wins[1].id, 'right');
      });

      const origRatio = await page.evaluate(() => {
        return app._test.dockContainers[0].root.ratio;
      });

      const splitter = page.locator('.dock-splitter').first();
      const box = await splitter.boundingBox();

      await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
      await page.mouse.down({ button: 'right' });
      await page.mouse.move(box.x + box.width / 2 + 80, box.y + box.height / 2);
      await page.mouse.up({ button: 'right' });

      const newRatio = await page.evaluate(() => {
        return app._test.dockContainers[0].root.ratio;
      });

      expect(newRatio).toBe(origRatio);
    });

    test('right-click does NOT trigger tab drag', async ({ page }) => {
      await executeSQL(page, "SELECT * INTO [t2] FROM [sample1]");
      await waitForWindow(page, 't2');

      await page.evaluate(() => {
        const wins = app._test.windows;
        app._test.mergeWindowsIntoTabs(wins[0].id, wins[1].id);
      });

      // Right-click on a dock tab should not initiate drag
      const tab = page.locator('.dock-tab').first();
      const box = await tab.boundingBox();

      await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
      await page.mouse.down({ button: 'right' });
      await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2 + 50);
      await page.mouse.up({ button: 'right' });

      // Windows should still be docked (no undocking happened)
      const result = await page.evaluate(() => ({
        dockCount: app._test.dockContainers.length,
        allDocked: app._test.windows.every(w => w.dockId !== null),
      }));
      expect(result.dockCount).toBe(1);
      expect(result.allDocked).toBe(true);
    });

    test('right-click does NOT trigger column resize', async ({ page }) => {
      // Wait for table to render with column widths
      await page.waitForSelector('.subwindow th .col-resize-handle');

      const origWidths = await page.evaluate(() => {
        const win = app._test.windows[0];
        return win.colWidths ? [...win.colWidths] : null;
      });

      const handle = page.locator('.col-resize-handle').first();
      const box = await handle.boundingBox();

      await page.mouse.move(box.x + 1, box.y + box.height / 2);
      await page.mouse.down({ button: 'right' });
      await page.mouse.move(box.x + 80, box.y + box.height / 2);
      await page.mouse.up({ button: 'right' });

      const newWidths = await page.evaluate(() => {
        const win = app._test.windows[0];
        return win.colWidths ? [...win.colWidths] : null;
      });

      expect(newWidths).toEqual(origWidths);
    });

    test('right-click does NOT trigger console resize', async ({ page }) => {
      const origHeight = await page.evaluate(() => {
        return document.getElementById('console-panel').offsetHeight;
      });

      const handle = page.locator('#console-resize-handle');
      const box = await handle.boundingBox();

      await page.mouse.move(box.x + box.width / 2, box.y + 1);
      await page.mouse.down({ button: 'right' });
      await page.mouse.move(box.x + box.width / 2, box.y - 80);
      await page.mouse.up({ button: 'right' });

      const newHeight = await page.evaluate(() => {
        return document.getElementById('console-panel').offsetHeight;
      });

      expect(newHeight).toBe(origHeight);
    });

    test('right-click on data cell does NOT trigger cell selection via focusin', async ({ page }) => {
      await page.waitForSelector('.subwindow td.data-cell');

      // Clear any existing selection
      await page.evaluate(() => {
        const win = app._test.windows[0];
        win.selectedCells = new Set();
        win.anchorCell = null;
      });

      const cell = page.locator('.subwindow td.data-cell').first();
      const box = await cell.boundingBox();

      await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2, { button: 'right' });

      const result = await page.evaluate(() => {
        const win = app._test.windows[0];
        return {
          selectedCount: win.selectedCells ? win.selectedCells.size : 0,
          hasAnchor: win.anchorCell !== null,
        };
      });

      expect(result.selectedCount).toBe(0);
      expect(result.hasAnchor).toBe(false);
    });
  });

  // =========================================================================
  // 2. Ghost mode for shift-drag
  // =========================================================================
  test.describe('Ghost mode for shift-drag', () => {

    test.beforeEach(async ({ page }) => {
      await openApp(page);
      await uploadFile(page, 'sample1.csv');
      await waitForWindow(page, 'sample1');
    });

    test('shift-drag creates ghost element with correct properties', async ({ page }) => {
      const titlebar = page.locator('.subwindow .win-titlebar').first();
      const box = await titlebar.boundingBox();

      const origPos = await page.evaluate(() => {
        const win = app._test.windows[0];
        return { left: win.el.style.left, top: win.el.style.top };
      });

      // Start shift-drag
      await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
      await page.keyboard.down('Shift');
      await page.mouse.down({ button: 'left' });
      await page.mouse.move(box.x + box.width / 2 + 30, box.y + box.height / 2 + 30);

      // Check ghost exists with correct styling
      const ghostInfo = await page.evaluate(() => {
        const ghost = document.querySelector('body > .subwindow[style*="position: fixed"]');
        if (!ghost) return null;
        return {
          opacity: ghost.style.opacity,
          pointerEvents: ghost.style.pointerEvents,
          zIndex: ghost.style.zIndex,
          hasTitle: ghost.querySelector('.win-title') !== null,
        };
      });

      expect(ghostInfo).not.toBeNull();
      expect(ghostInfo.opacity).toBe('0.7');
      expect(ghostInfo.pointerEvents).toBe('none');
      expect(ghostInfo.zIndex).toBe('999999');
      expect(ghostInfo.hasTitle).toBe(true);

      // Original window stays in place
      const curPos = await page.evaluate(() => {
        const win = app._test.windows[0];
        return { left: win.el.style.left, top: win.el.style.top };
      });
      expect(curPos.left).toBe(origPos.left);
      expect(curPos.top).toBe(origPos.top);

      await page.mouse.up({ button: 'left' });
      await page.keyboard.up('Shift');
    });

    test('drag without shift moves window normally (no ghost)', async ({ page }) => {
      const titlebar = page.locator('.subwindow .win-titlebar').first();
      const box = await titlebar.boundingBox();

      const origPos = await page.evaluate(() => {
        const win = app._test.windows[0];
        return { left: parseInt(win.el.style.left), top: parseInt(win.el.style.top) };
      });

      // Normal drag without shift
      await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
      await page.mouse.down({ button: 'left' });
      await page.mouse.move(box.x + box.width / 2 + 40, box.y + box.height / 2 + 40);

      // No ghost should exist
      const ghostExists = await page.evaluate(() => {
        return document.querySelector('body > .subwindow[style*="position: fixed"]') !== null;
      });
      expect(ghostExists).toBe(false);

      // Window should have moved
      const newPos = await page.evaluate(() => {
        const win = app._test.windows[0];
        return { left: parseInt(win.el.style.left), top: parseInt(win.el.style.top) };
      });
      expect(newPos.left).not.toBe(origPos.left);
      expect(newPos.top).not.toBe(origPos.top);

      await page.mouse.up({ button: 'left' });
    });
  });

  // =========================================================================
  // 3. Shift only needed to initiate ghost mode
  // =========================================================================
  test.describe('Shift only needed to initiate', () => {

    test.beforeEach(async ({ page }) => {
      await openApp(page);
      await uploadFile(page, 'sample1.csv');
      await waitForWindow(page, 'sample1');
      await executeSQL(page, "SELECT * INTO [t2] FROM [sample1]");
      await waitForWindow(page, 't2');
    });

    test('releasing shift keeps ghost mode active', async ({ page }) => {
      // Position windows so they don't overlap for clear drop targeting
      await page.evaluate(() => {
        const wins = app._test.windows;
        wins[0].el.style.left = '10px';
        wins[0].el.style.top = '10px';
        wins[0].el.style.width = '300px';
        wins[0].el.style.height = '250px';
        wins[1].el.style.left = '320px';
        wins[1].el.style.top = '10px';
        wins[1].el.style.width = '300px';
        wins[1].el.style.height = '250px';
      });

      const titlebar = page.locator('.subwindow .win-titlebar').first();
      const box = await titlebar.boundingBox();

      // Initiate ghost with shift held
      await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
      await page.keyboard.down('Shift');
      await page.mouse.down({ button: 'left' });
      await page.mouse.move(box.x + box.width / 2 + 20, box.y + box.height / 2 + 20);

      // Verify ghost exists
      let ghostExists = await page.evaluate(() => {
        return document.querySelector('body > .subwindow[style*="position: fixed"]') !== null;
      });
      expect(ghostExists).toBe(true);

      // Release shift while still dragging
      await page.keyboard.up('Shift');
      await page.mouse.move(box.x + box.width / 2 + 50, box.y + box.height / 2 + 50);

      // Ghost should still exist
      ghostExists = await page.evaluate(() => {
        return document.querySelector('body > .subwindow[style*="position: fixed"]') !== null;
      });
      expect(ghostExists).toBe(true);

      await page.mouse.up({ button: 'left' });
    });

    test('drop targets are still detected after shift is released', async ({ page }) => {
      // Position windows clearly apart
      await page.evaluate(() => {
        const wins = app._test.windows;
        wins[0].el.style.left = '10px';
        wins[0].el.style.top = '10px';
        wins[0].el.style.width = '300px';
        wins[0].el.style.height = '300px';
        wins[1].el.style.left = '320px';
        wins[1].el.style.top = '10px';
        wins[1].el.style.width = '300px';
        wins[1].el.style.height = '300px';
      });

      const firstTitlebar = page.locator('.subwindow .win-titlebar').first();
      const firstBox = await firstTitlebar.boundingBox();

      // Get second window's titlebar position for drop target
      const secondWin = page.locator('.subwindow').nth(1);
      const secondBox = await secondWin.boundingBox();

      // Start shift-drag from first window
      await page.mouse.move(firstBox.x + firstBox.width / 2, firstBox.y + firstBox.height / 2);
      await page.keyboard.down('Shift');
      await page.mouse.down({ button: 'left' });
      await page.mouse.move(firstBox.x + firstBox.width / 2 + 20, firstBox.y + firstBox.height / 2);

      // Release shift
      await page.keyboard.up('Shift');

      // Move over second window's titlebar (center zone = tab)
      await page.mouse.move(secondBox.x + secondBox.width / 2, secondBox.y + 10);

      // Drop overlay should be showing
      const hasOverlay = await page.evaluate(() => {
        return document.querySelector('.dock-drop-overlay') !== null;
      });
      expect(hasOverlay).toBe(true);

      // Drop - should create a dock
      await page.mouse.up({ button: 'left' });

      const result = await page.evaluate(() => ({
        dockCount: app._test.dockContainers.length,
      }));
      expect(result.dockCount).toBe(1);
    });
  });

  // =========================================================================
  // 4. Single-tab hidden tab bar
  // =========================================================================
  test.describe('Single-tab hidden tab bar', () => {

    test.beforeEach(async ({ page }) => {
      await openApp(page);
      await uploadFile(page, 'sample1.csv');
      await waitForWindow(page, 'sample1');
      await executeSQL(page, "SELECT * INTO [t2] FROM [sample1]");
      await waitForWindow(page, 't2');
    });

    test('vertical split with single tabs gets single-tab class on non-controls tab bars', async ({ page }) => {
      await page.evaluate(() => {
        const wins = app._test.windows;
        app._test.mergeWindowsAsSplit(wins[0].id, wins[1].id, 'bottom');
      });

      const result = await page.evaluate(() => {
        const dock = app._test.dockContainers[0];
        const tabBars = dock.el.querySelectorAll('.dock-tab-bar');
        const singleTabBars = dock.el.querySelectorAll('.dock-tab-bar.single-tab');
        const leafCount = dock.el.querySelectorAll('.dock-leaf').length;
        // The tab bar that is on .dock-root direct child (controls tab bar) should NOT have single-tab
        const rootTabBar = dock.el.querySelector('.dock-root > .dock-leaf > .dock-tab-bar, .dock-root > .dock-split .dock-leaf > .dock-tab-bar');
        return {
          totalTabBars: tabBars.length,
          singleTabCount: singleTabBars.length,
          leafCount,
        };
      });

      // In a split, both leaves have parent (the split node), so both should get single-tab
      expect(result.leafCount).toBe(2);
      expect(result.singleTabCount).toBe(2);
    });

    test('single-tab tab bar still contains a tab element', async ({ page }) => {
      await page.evaluate(() => {
        const wins = app._test.windows;
        app._test.mergeWindowsAsSplit(wins[0].id, wins[1].id, 'bottom');
      });

      const result = await page.evaluate(() => {
        const singleTabBar = document.querySelector('.dock-tab-bar.single-tab');
        if (!singleTabBar) return null;
        const tabs = singleTabBar.querySelectorAll('.dock-tab');
        const tab = tabs[0];
        return {
          tabCount: tabs.length,
          hasTitle: tab ? tab.querySelector('.dock-tab-title') !== null : false,
          tabFlexStyle: tab ? window.getComputedStyle(tab).flex : '',
        };
      });

      expect(result).not.toBeNull();
      expect(result.tabCount).toBe(1);
      expect(result.hasTitle).toBe(true);
    });

    test('controls tab bar (root > leaf direct child) is always visible', async ({ page }) => {
      // A single leaf dock (tabs only, no split) should NOT have single-tab
      await page.evaluate(() => {
        const wins = app._test.windows;
        app._test.mergeWindowsIntoTabs(wins[0].id, wins[1].id);
      });

      const result = await page.evaluate(() => {
        const dock = app._test.dockContainers[0];
        // root is a leaf directly (no parent), so its tab bar should NOT get single-tab
        const tabBar = dock.root.tabBarEl;
        return {
          hasSingleTab: tabBar.classList.contains('single-tab'),
          tabCount: tabBar.querySelectorAll('.dock-tab').length,
          isVisible: tabBar.offsetParent !== null,
        };
      });

      expect(result.hasSingleTab).toBe(false);
      expect(result.tabCount).toBe(2);
      expect(result.isVisible).toBe(true);
    });

    test('single-tab class is removed when a second tab is added', async ({ page }) => {
      await page.evaluate(() => {
        const wins = app._test.windows;
        app._test.mergeWindowsAsSplit(wins[0].id, wins[1].id, 'right');
      });

      // Initially both leaves have single-tab
      let singleTabCount = await page.evaluate(() => {
        return document.querySelectorAll('.dock-tab-bar.single-tab').length;
      });
      expect(singleTabCount).toBe(2);

      // Add a third window as a tab to one leaf
      await executeSQL(page, "SELECT * INTO [t3] FROM [sample1]");
      await waitForWindow(page, 't3');

      await page.evaluate(() => {
        const wins = app._test.windows;
        const dock = app._test.dockContainers[0];
        // Find a leaf to tab into
        function findLeaves(node) {
          if (node.type === 'leaf') return [node];
          return node.children.flatMap(c => findLeaves(c));
        }
        const leaves = findLeaves(dock.root);
        app._test.dockWindowAsTab(wins[2].id, leaves[0], dock);
        wins[2].el.classList.add('docked');
      });

      singleTabCount = await page.evaluate(() => {
        return document.querySelectorAll('.dock-tab-bar.single-tab').length;
      });
      // Only one leaf still has a single tab
      expect(singleTabCount).toBe(1);
    });
  });

  // =========================================================================
  // 5. Self-drop is no-op
  // =========================================================================
  test.describe('Self-drop is no-op', () => {

    test.beforeEach(async ({ page }) => {
      await openApp(page);
      await uploadFile(page, 'sample1.csv');
      await waitForWindow(page, 'sample1');
      await executeSQL(page, "SELECT * INTO [t2] FROM [sample1]");
      await waitForWindow(page, 't2');
    });

    test('dropping tab on same leaf center zone is a no-op', async ({ page }) => {
      await page.evaluate(() => {
        const wins = app._test.windows;
        app._test.mergeWindowsIntoTabs(wins[0].id, wins[1].id);
      });

      const beforeState = await page.evaluate(() => ({
        dockCount: app._test.dockContainers.length,
        tabCount: document.querySelectorAll('.dock-tab').length,
        windowCount: app._test.windows.length,
      }));

      // Get tab position
      const tab = page.locator('.dock-tab').first();
      const tabBox = await tab.boundingBox();

      // Shift-drag the tab and drop on same leaf center
      await page.mouse.move(tabBox.x + tabBox.width / 2, tabBox.y + tabBox.height / 2);
      await page.keyboard.down('Shift');
      await page.mouse.down({ button: 'left' });
      // Move past threshold to initiate undock mode
      await page.mouse.move(tabBox.x + tabBox.width / 2 + 10, tabBox.y + tabBox.height / 2 + 40);
      // Move back to the same leaf center area (the dock body)
      const dockBody = page.locator('.dock-leaf-content').first();
      const bodyBox = await dockBody.boundingBox();
      await page.mouse.move(bodyBox.x + bodyBox.width / 2, bodyBox.y + bodyBox.height / 2);
      await page.mouse.up({ button: 'left' });
      await page.keyboard.up('Shift');

      const afterState = await page.evaluate(() => ({
        dockCount: app._test.dockContainers.length,
        tabCount: document.querySelectorAll('.dock-tab').length,
        windowCount: app._test.windows.length,
      }));

      expect(afterState.dockCount).toBe(beforeState.dockCount);
      expect(afterState.tabCount).toBe(beforeState.tabCount);
      expect(afterState.windowCount).toBe(beforeState.windowCount);
    });

    test('dropping tab with cursor still over source dock does NOT undock', async ({ page }) => {
      await page.evaluate(() => {
        const wins = app._test.windows;
        app._test.mergeWindowsIntoTabs(wins[0].id, wins[1].id);
      });

      const beforeState = await page.evaluate(() => ({
        dockCount: app._test.dockContainers.length,
        allDocked: app._test.windows.every(w => w.dockId !== null),
      }));

      // Get tab position
      const tab = page.locator('.dock-tab').first();
      const tabBox = await tab.boundingBox();

      // Shift-drag the tab downward but stay within dock container bounds
      const dockContainer = page.locator('.dock-container');
      const dockBox = await dockContainer.boundingBox();

      await page.mouse.move(tabBox.x + tabBox.width / 2, tabBox.y + tabBox.height / 2);
      await page.keyboard.down('Shift');
      await page.mouse.down({ button: 'left' });
      // Move past threshold but stay within dock container
      await page.mouse.move(tabBox.x + tabBox.width / 2, tabBox.y + tabBox.height / 2 + 40);
      // Stay over the dock container body (no outside target, no valid drop)
      await page.mouse.move(dockBox.x + dockBox.width / 2, dockBox.y + dockBox.height / 2);
      await page.mouse.up({ button: 'left' });
      await page.keyboard.up('Shift');

      const afterState = await page.evaluate(() => ({
        dockCount: app._test.dockContainers.length,
        allDocked: app._test.windows.every(w => w.dockId !== null),
      }));

      expect(afterState.dockCount).toBe(beforeState.dockCount);
      expect(afterState.allDocked).toBe(true);
    });
  });

  // =========================================================================
  // 6. Undock retains size
  // =========================================================================
  test.describe('Undock retains size', () => {

    test.beforeEach(async ({ page }) => {
      await openApp(page);
      await uploadFile(page, 'sample1.csv');
      await waitForWindow(page, 'sample1');
      await executeSQL(page, "SELECT * INTO [t2] FROM [sample1]");
      await waitForWindow(page, 't2');
    });

    test('undocked window has dimensions matching the leaf pane', async ({ page }) => {
      // Create a 3-window dock so undocking one does not dissolve the entire dock.
      // With only 2 windows in a split, undocking one dissolves the dock and both
      // become standalone, making it ambiguous which window we actually dragged.
      await executeSQL(page, "SELECT * INTO [t3] FROM [sample1]");
      await waitForWindow(page, 't3');

      await page.evaluate(() => {
        const wins = app._test.windows;
        // Tab first two windows, then split the third to the right
        app._test.mergeWindowsIntoTabs(wins[0].id, wins[1].id);
        const dock = app._test.dockContainers[0];
        app._test.dockWindowAsSplit(wins[2].id, dock.root, dock, 'right');
        wins[2].el.classList.add('docked');
        // Set the dock to a known size
        dock.el.style.width = '600px';
        dock.el.style.height = '400px';
      });
      await page.waitForTimeout(100);

      // Record the ID and leaf dimensions of the window we will undock
      const info = await page.evaluate(() => {
        const win = app._test.windows[0];
        const leaf = win.dockLeaf;
        const rect = leaf.el.getBoundingClientRect();
        return {
          winId: win.id,
          leafWidth: Math.round(rect.width),
          leafHeight: Math.round(rect.height),
        };
      });

      // Undock via the tab drag code path. The setupTabDragFromMousedown function
      // records leaf dimensions and applies them to the undocked window.
      // We simulate the full shift-drag sequence using dispatchEvent to ensure
      // shiftKey is properly set on mousemove events.
      const undockedSize = await page.evaluate((targetId) => {
        const win = app._test.windows.find(w => w.id === targetId);
        if (!win || !win.dockLeaf) return { error: 'win not found or not docked' };

        const tabEl = win.dockLeaf.tabBarEl.querySelector(`.dock-tab[data-win-id="${targetId}"]`);
        if (!tabEl) return { error: 'tab not found' };

        const dock = app._test.dockContainers.find(d => d.id === win.dockId);
        const dockRect = dock.el.getBoundingClientRect();
        const tabRect = tabEl.getBoundingClientRect();
        const startX = tabRect.left + tabRect.width / 2;
        const startY = tabRect.top + tabRect.height / 2;

        // Mousedown on the tab (triggers setupTabDragFromMousedown)
        tabEl.dispatchEvent(new MouseEvent('mousedown', {
          button: 0, clientX: startX, clientY: startY, bubbles: true
        }));

        // Mousemove with shiftKey to enter undock mode (>5px threshold)
        document.dispatchEvent(new MouseEvent('mousemove', {
          clientX: startX, clientY: startY + 10, shiftKey: true, bubbles: true
        }));

        // Move far outside the dock container
        const farX = dockRect.right + 150;
        const farY = dockRect.bottom + 150;
        document.dispatchEvent(new MouseEvent('mousemove', {
          clientX: farX, clientY: farY, shiftKey: true, bubbles: true
        }));

        // Mouseup outside the dock
        document.dispatchEvent(new MouseEvent('mouseup', {
          clientX: farX, clientY: farY, bubbles: true
        }));

        return {
          dockId: win.dockId,
          width: parseInt(win.el.style.width),
          height: parseInt(win.el.style.height),
        };
      }, info.winId);

      expect(undockedSize.dockId).toBeNull();
      expect(undockedSize.width).toBe(info.leafWidth);
      expect(undockedSize.height).toBe(info.leafHeight);
    });

    test('undocked window size is not 60% of dock container', async ({ page }) => {
      await page.evaluate(() => {
        const wins = app._test.windows;
        app._test.mergeWindowsAsSplit(wins[0].id, wins[1].id, 'right');
        const dock = app._test.dockContainers[0];
        dock.el.style.width = '800px';
        dock.el.style.height = '500px';
      });
      await page.waitForTimeout(100);

      const dockSize = await page.evaluate(() => {
        const dock = app._test.dockContainers[0];
        return {
          width: parseInt(dock.el.style.width),
          height: parseInt(dock.el.style.height),
        };
      });

      // Undock via shift-drag outside
      const dockEl = page.locator('.dock-container');
      const dockBox = await dockEl.boundingBox();
      const tab = page.locator('.dock-tab').first();
      const tabBox = await tab.boundingBox();

      await page.mouse.move(tabBox.x + tabBox.width / 2, tabBox.y + tabBox.height / 2);
      await page.keyboard.down('Shift');
      await page.mouse.down({ button: 'left' });
      await page.mouse.move(dockBox.x + dockBox.width + 100, dockBox.y + dockBox.height + 100);
      await page.mouse.up({ button: 'left' });
      await page.keyboard.up('Shift');

      const undockedSize = await page.evaluate(() => {
        const win = app._test.windows.find(w => w.dockId === null);
        if (!win) return null;
        return {
          width: parseInt(win.el.style.width),
          height: parseInt(win.el.style.height),
        };
      });

      // Size should NOT be 60% of dock (which would be old/broken behavior)
      expect(undockedSize.width).not.toBe(Math.round(dockSize.width * 0.6));
      expect(undockedSize.height).not.toBe(Math.round(dockSize.height * 0.6));
    });
  });

  // =========================================================================
  // 7. Diagonal drop zones
  // =========================================================================
  test.describe('Diagonal drop zones', () => {

    test.beforeEach(async ({ page }) => {
      await openApp(page);
    });

    test('titlebar area returns center zone', async ({ page }) => {
      const zone = await page.evaluate(() => {
        // Simulate a rect at (100, 100) with width 400, height 300
        const rect = { left: 100, top: 100, width: 400, height: 300, right: 500, bottom: 400 };
        // Point in titlebar area (within first 32px)
        return app._test.getDropZone(300, 115, rect, 32);
      });
      expect(zone).toBe('center');
    });

    test('top triangle returns top zone', async ({ page }) => {
      const zone = await page.evaluate(() => {
        // rect at (0, 0), width 400, height 332 (300 body + 32 titlebar)
        const rect = { left: 0, top: 0, width: 400, height: 332, right: 400, bottom: 332 };
        // Body starts at y=32. In body coords: center is (200, 150)
        // Top triangle: point at body center-x, just below titlebar
        // bx=200, by=10 -> u=0.5, v=0.033 -> v < u (0.033 < 0.5) and v < 1-u (0.033 < 0.5)
        return app._test.getDropZone(200, 42, rect, 32);
      });
      expect(zone).toBe('top');
    });

    test('bottom triangle returns bottom zone', async ({ page }) => {
      const zone = await page.evaluate(() => {
        const rect = { left: 0, top: 0, width: 400, height: 332, right: 400, bottom: 332 };
        // Body: bw=400, bh=300. Point near bottom center: bx=200, by=290
        // u=0.5, v=290/300=0.967 -> v > u (0.967 > 0.5) and v > 1-u (0.967 > 0.5)
        return app._test.getDropZone(200, 322, rect, 32);
      });
      expect(zone).toBe('bottom');
    });

    test('left triangle returns left zone', async ({ page }) => {
      const zone = await page.evaluate(() => {
        const rect = { left: 0, top: 0, width: 400, height: 332, right: 400, bottom: 332 };
        // Body: bw=400, bh=300. Left triangle: bx=20, by=150 (center height)
        // u=20/400=0.05, v=150/300=0.5 -> v >= u (0.5 >= 0.05) and v <= 1-u (0.5 <= 0.95)
        return app._test.getDropZone(20, 182, rect, 32);
      });
      expect(zone).toBe('left');
    });

    test('right triangle returns right zone', async ({ page }) => {
      const zone = await page.evaluate(() => {
        const rect = { left: 0, top: 0, width: 400, height: 332, right: 400, bottom: 332 };
        // Body: bw=400, bh=300. Right triangle: bx=380, by=150
        // u=380/400=0.95, v=150/300=0.5 -> NOT (v < u && v < 1-u), NOT (v > u && v > 1-u)
        // NOT (v >= u && v <= 1-u) -> 0.5 >= 0.95 is false -> falls to 'right'
        return app._test.getDropZone(380, 182, rect, 32);
      });
      expect(zone).toBe('right');
    });

    test('exact center of body is classified by diagonal math', async ({ page }) => {
      const zone = await page.evaluate(() => {
        const rect = { left: 0, top: 0, width: 400, height: 332, right: 400, bottom: 332 };
        // Exact center of body: bx=200, by=150 -> u=0.5, v=0.5
        // v < u? 0.5 < 0.5 -> false
        // v > u? 0.5 > 0.5 -> false
        // v >= u && v <= 1-u? 0.5 >= 0.5 && 0.5 <= 0.5 -> true -> 'left'
        return app._test.getDropZone(200, 182, rect, 32);
      });
      expect(zone).toBe('left');
    });

    test('top-right corner uses diagonal classification', async ({ page }) => {
      const zone = await page.evaluate(() => {
        const rect = { left: 0, top: 0, width: 400, height: 332, right: 400, bottom: 332 };
        // bx=380, by=10 -> u=0.95, v=0.033
        // v < u? 0.033 < 0.95 -> true; v < 1-u? 0.033 < 0.05 -> true -> 'top'
        return app._test.getDropZone(380, 42, rect, 32);
      });
      expect(zone).toBe('top');
    });

    test('bottom-left corner uses diagonal classification', async ({ page }) => {
      const zone = await page.evaluate(() => {
        const rect = { left: 0, top: 0, width: 400, height: 332, right: 400, bottom: 332 };
        // bx=20, by=290 -> u=0.05, v=0.967
        // v > u? 0.967 > 0.05 -> true; v > 1-u? 0.967 > 0.95 -> true -> 'bottom'
        return app._test.getDropZone(20, 322, rect, 32);
      });
      expect(zone).toBe('bottom');
    });

    test('custom titlebar height is respected', async ({ page }) => {
      const zone = await page.evaluate(() => {
        const rect = { left: 0, top: 0, width: 400, height: 350, right: 400, bottom: 350 };
        // titlebar height = 50. Point at y=40 (within titlebar)
        return app._test.getDropZone(200, 40, rect, 50);
      });
      expect(zone).toBe('center');
    });

    test('zero-height body defaults to center', async ({ page }) => {
      const zone = await page.evaluate(() => {
        // rect where height equals titlebar height (no body)
        const rect = { left: 0, top: 0, width: 400, height: 32, right: 400, bottom: 32 };
        return app._test.getDropZone(200, 33, rect, 32);
      });
      // bh = 32-32 = 0, so returns 'center'
      expect(zone).toBe('center');
    });
  });

  // =========================================================================
  // 8. Drop overlay is rectangular
  // =========================================================================
  test.describe('Drop overlay is rectangular', () => {

    test.beforeEach(async ({ page }) => {
      await openApp(page);
      await uploadFile(page, 'sample1.csv');
      await waitForWindow(page, 'sample1');
      await executeSQL(page, "SELECT * INTO [t2] FROM [sample1]");
      await waitForWindow(page, 't2');
      // Position windows apart
      await page.evaluate(() => {
        const wins = app._test.windows;
        wins[0].el.style.left = '10px';
        wins[0].el.style.top = '10px';
        wins[0].el.style.width = '300px';
        wins[0].el.style.height = '300px';
        wins[1].el.style.left = '320px';
        wins[1].el.style.top = '10px';
        wins[1].el.style.width = '300px';
        wins[1].el.style.height = '300px';
      });
    });

    test('center zone overlay covers entire element', async ({ page }) => {
      const firstTitlebar = page.locator('.subwindow .win-titlebar').first();
      const firstBox = await firstTitlebar.boundingBox();

      const secondWin = page.locator('.subwindow').nth(1);
      const secondBox = await secondWin.boundingBox();

      // Shift-drag from first window to second window's titlebar (center zone)
      await page.mouse.move(firstBox.x + firstBox.width / 2, firstBox.y + firstBox.height / 2);
      await page.keyboard.down('Shift');
      await page.mouse.down({ button: 'left' });
      await page.mouse.move(secondBox.x + secondBox.width / 2, secondBox.y + 10);

      const overlay = await page.evaluate(() => {
        const o = document.querySelector('.dock-drop-overlay');
        if (!o) return null;
        const zone = o.querySelector('.dock-drop-zone');
        if (!zone) return null;
        return {
          zoneInset: zone.style.inset,
        };
      });

      expect(overlay).not.toBeNull();
      expect(overlay.zoneInset).toBe('0px');

      await page.mouse.up({ button: 'left' });
      await page.keyboard.up('Shift');
    });

    test('top zone overlay covers top half of body area', async ({ page }) => {
      const firstTitlebar = page.locator('.subwindow .win-titlebar').first();
      const firstBox = await firstTitlebar.boundingBox();

      const secondWin = page.locator('.subwindow').nth(1);
      const secondBox = await secondWin.boundingBox();

      // Shift-drag to top zone of second window (center x, just below titlebar)
      await page.mouse.move(firstBox.x + firstBox.width / 2, firstBox.y + firstBox.height / 2);
      await page.keyboard.down('Shift');
      await page.mouse.down({ button: 'left' });
      // Drop in the top triangle: center x, about 1/4 from the top of body
      await page.mouse.move(secondBox.x + secondBox.width / 2, secondBox.y + 50);

      const overlay = await page.evaluate(() => {
        const o = document.querySelector('.dock-drop-overlay');
        if (!o) return null;
        const zone = o.querySelector('.dock-drop-zone');
        if (!zone) return null;
        const cs = zone.style;
        return {
          left: cs.left,
          right: cs.right,
          top: cs.top,
          height: cs.height,
        };
      });

      expect(overlay).not.toBeNull();
      expect(overlay.left).toBe('0px');
      expect(overlay.right).toBe('0px');
      // top should be titlebar height (32px)
      expect(parseInt(overlay.top)).toBe(32);
      // height should be about half the body area
      const bodyH = 300 - 32;
      expect(parseInt(overlay.height)).toBe(Math.floor(bodyH / 2));

      await page.mouse.up({ button: 'left' });
      await page.keyboard.up('Shift');
    });

    test('bottom zone overlay covers bottom half of body area', async ({ page }) => {
      const firstTitlebar = page.locator('.subwindow .win-titlebar').first();
      const firstBox = await firstTitlebar.boundingBox();

      const secondWin = page.locator('.subwindow').nth(1);
      const secondBox = await secondWin.boundingBox();

      // Shift-drag to bottom zone of second window
      await page.mouse.move(firstBox.x + firstBox.width / 2, firstBox.y + firstBox.height / 2);
      await page.keyboard.down('Shift');
      await page.mouse.down({ button: 'left' });
      // Bottom triangle: center x, near bottom
      await page.mouse.move(secondBox.x + secondBox.width / 2, secondBox.y + secondBox.height - 20);

      const overlay = await page.evaluate(() => {
        const o = document.querySelector('.dock-drop-overlay');
        if (!o) return null;
        const zone = o.querySelector('.dock-drop-zone');
        if (!zone) return null;
        const cs = zone.style;
        return {
          left: cs.left,
          right: cs.right,
          bottom: cs.bottom,
          height: cs.height,
        };
      });

      expect(overlay).not.toBeNull();
      expect(overlay.left).toBe('0px');
      expect(overlay.right).toBe('0px');
      expect(overlay.bottom).toBe('0px');
      const bodyH = 300 - 32;
      expect(parseInt(overlay.height)).toBe(Math.floor(bodyH / 2));

      await page.mouse.up({ button: 'left' });
      await page.keyboard.up('Shift');
    });

    test('left zone overlay covers left half of body area', async ({ page }) => {
      const firstTitlebar = page.locator('.subwindow .win-titlebar').first();
      const firstBox = await firstTitlebar.boundingBox();

      const secondWin = page.locator('.subwindow').nth(1);
      const secondBox = await secondWin.boundingBox();

      // Shift-drag to left zone: left edge, vertical center of body
      await page.mouse.move(firstBox.x + firstBox.width / 2, firstBox.y + firstBox.height / 2);
      await page.keyboard.down('Shift');
      await page.mouse.down({ button: 'left' });
      await page.mouse.move(secondBox.x + 15, secondBox.y + secondBox.height / 2);

      const overlay = await page.evaluate(() => {
        const o = document.querySelector('.dock-drop-overlay');
        if (!o) return null;
        const zone = o.querySelector('.dock-drop-zone');
        if (!zone) return null;
        const cs = zone.style;
        return {
          top: cs.top,
          bottom: cs.bottom,
          left: cs.left,
          width: cs.width,
        };
      });

      expect(overlay).not.toBeNull();
      expect(parseInt(overlay.top)).toBe(32);
      expect(overlay.bottom).toBe('0px');
      expect(overlay.left).toBe('0px');
      expect(parseInt(overlay.width)).toBe(Math.floor(300 / 2));

      await page.mouse.up({ button: 'left' });
      await page.keyboard.up('Shift');
    });

    test('right zone overlay covers right half of body area', async ({ page }) => {
      const firstTitlebar = page.locator('.subwindow .win-titlebar').first();
      const firstBox = await firstTitlebar.boundingBox();

      const secondWin = page.locator('.subwindow').nth(1);
      const secondBox = await secondWin.boundingBox();

      // Shift-drag to right zone: right edge, vertical center of body
      await page.mouse.move(firstBox.x + firstBox.width / 2, firstBox.y + firstBox.height / 2);
      await page.keyboard.down('Shift');
      await page.mouse.down({ button: 'left' });
      await page.mouse.move(secondBox.x + secondBox.width - 15, secondBox.y + secondBox.height / 2);

      const overlay = await page.evaluate(() => {
        const o = document.querySelector('.dock-drop-overlay');
        if (!o) return null;
        const zone = o.querySelector('.dock-drop-zone');
        if (!zone) return null;
        const cs = zone.style;
        return {
          top: cs.top,
          bottom: cs.bottom,
          right: cs.right,
          width: cs.width,
        };
      });

      expect(overlay).not.toBeNull();
      expect(parseInt(overlay.top)).toBe(32);
      expect(overlay.bottom).toBe('0px');
      expect(overlay.right).toBe('0px');
      expect(parseInt(overlay.width)).toBe(Math.floor(300 / 2));

      await page.mouse.up({ button: 'left' });
      await page.keyboard.up('Shift');
    });
  });

  // =========================================================================
  // 9. Double-click tab maximizes
  // =========================================================================
  test.describe('Double-click tab maximizes', () => {

    test.beforeEach(async ({ page }) => {
      await openApp(page);
      await uploadFile(page, 'sample1.csv');
      await waitForWindow(page, 'sample1');
      await executeSQL(page, "SELECT * INTO [t2] FROM [sample1]");
      await waitForWindow(page, 't2');
    });

    test('double-clicking a dock tab toggles maximize on dock container', async ({ page }) => {
      await page.evaluate(() => {
        const wins = app._test.windows;
        app._test.mergeWindowsIntoTabs(wins[0].id, wins[1].id);
        const dock = app._test.dockContainers[0];
        dock.el.style.left = '50px';
        dock.el.style.top = '50px';
        dock.el.style.width = '400px';
        dock.el.style.height = '300px';
      });

      const beforeMax = await page.evaluate(() => {
        return app._test.dockContainers[0].maximized || false;
      });
      expect(beforeMax).toBe(false);

      // Double-click a dock tab
      const tab = page.locator('.dock-tab').first();
      await tab.dblclick();

      const afterFirstClick = await page.evaluate(() => {
        const dock = app._test.dockContainers[0];
        return {
          maximized: dock.maximized,
          left: dock.el.style.left,
          top: dock.el.style.top,
        };
      });
      expect(afterFirstClick.maximized).toBe(true);
      expect(afterFirstClick.left).toBe('0px');
      expect(afterFirstClick.top).toBe('0px');

      // Double-click again to restore
      const tab2 = page.locator('.dock-tab').first();
      await tab2.dblclick();

      const afterSecondClick = await page.evaluate(() => {
        const dock = app._test.dockContainers[0];
        return {
          maximized: dock.maximized,
          left: dock.el.style.left,
          top: dock.el.style.top,
        };
      });
      expect(afterSecondClick.maximized).toBe(false);
      expect(afterSecondClick.left).toBe('50px');
      expect(afterSecondClick.top).toBe('50px');
    });

    test('double-clicking tab bar empty area also maximizes', async ({ page }) => {
      await page.evaluate(() => {
        const wins = app._test.windows;
        app._test.mergeWindowsIntoTabs(wins[0].id, wins[1].id);
        const dock = app._test.dockContainers[0];
        dock.el.style.left = '50px';
        dock.el.style.top = '50px';
        dock.el.style.width = '500px';
        dock.el.style.height = '300px';
      });

      // The tab bar dblclick handler fires when clicking on the tab bar itself
      // (not on a .dock-tab child). Since tabs use flex:1, they fill the width,
      // so we dispatch the dblclick event programmatically on the tab bar element.
      const result = await page.evaluate(() => {
        const dock = app._test.dockContainers[0];
        const tabBar = dock.root.tabBarEl;
        // Dispatch dblclick directly on the tab bar (not on a tab child)
        tabBar.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
        return dock.maximized || false;
      });
      expect(result).toBe(true);

      // Verify it toggles back
      const restored = await page.evaluate(() => {
        const dock = app._test.dockContainers[0];
        const tabBar = dock.root.tabBarEl;
        tabBar.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
        return dock.maximized || false;
      });
      expect(restored).toBe(false);
    });
  });

  // =========================================================================
  // 10. Resize handle precision (no overflow hidden on dock-container)
  // =========================================================================
  test.describe('Resize handle precision', () => {

    test.beforeEach(async ({ page }) => {
      await openApp(page);
      await uploadFile(page, 'sample1.csv');
      await waitForWindow(page, 'sample1');
      await executeSQL(page, "SELECT * INTO [t2] FROM [sample1]");
      await waitForWindow(page, 't2');
    });

    test('dock-container does NOT have overflow hidden', async ({ page }) => {
      await page.evaluate(() => {
        const wins = app._test.windows;
        app._test.mergeWindowsIntoTabs(wins[0].id, wins[1].id);
      });

      const overflow = await page.evaluate(() => {
        const container = document.querySelector('.dock-container');
        return window.getComputedStyle(container).overflow;
      });

      expect(overflow).not.toBe('hidden');
    });

    test('dock-root has overflow hidden', async ({ page }) => {
      await page.evaluate(() => {
        const wins = app._test.windows;
        app._test.mergeWindowsIntoTabs(wins[0].id, wins[1].id);
      });

      const overflow = await page.evaluate(() => {
        const root = document.querySelector('.dock-root');
        return window.getComputedStyle(root).overflow;
      });

      expect(overflow).toBe('hidden');
    });

    test('resize handles extend beyond dock-container border and are not clipped', async ({ page }) => {
      await page.evaluate(() => {
        const wins = app._test.windows;
        app._test.mergeWindowsIntoTabs(wins[0].id, wins[1].id);
        const dock = app._test.dockContainers[0];
        dock.el.style.left = '50px';
        dock.el.style.top = '50px';
        dock.el.style.width = '400px';
        dock.el.style.height = '300px';
      });

      // Check that resize handles exist and are visible (not clipped by overflow)
      const result = await page.evaluate(() => {
        const container = document.querySelector('.dock-container');
        const handles = container.querySelectorAll('[class*="rh-"]');
        const containerRect = container.getBoundingClientRect();
        let handleOutsideBorder = false;

        for (const h of handles) {
          const hRect = h.getBoundingClientRect();
          // At least one handle edge should extend at or beyond container edge
          if (hRect.right >= containerRect.right - 1 ||
              hRect.bottom >= containerRect.bottom - 1 ||
              hRect.left <= containerRect.left + 1 ||
              hRect.top <= containerRect.top + 1) {
            handleOutsideBorder = true;
            break;
          }
        }

        return {
          handleCount: handles.length,
          handleOutsideBorder,
          containerOverflow: window.getComputedStyle(container).overflow,
        };
      });

      expect(result.handleCount).toBeGreaterThan(0);
      expect(result.handleOutsideBorder).toBe(true);
      expect(result.containerOverflow).not.toBe('hidden');
    });

    test('dock-container resize handles respond to left-click', async ({ page }) => {
      await page.evaluate(() => {
        const wins = app._test.windows;
        app._test.mergeWindowsIntoTabs(wins[0].id, wins[1].id);
        const dock = app._test.dockContainers[0];
        dock.el.style.left = '50px';
        dock.el.style.top = '50px';
        dock.el.style.width = '400px';
        dock.el.style.height = '300px';
      });

      const origSize = await page.evaluate(() => {
        const dock = app._test.dockContainers[0];
        return { width: parseInt(dock.el.style.width), height: parseInt(dock.el.style.height) };
      });

      const handle = page.locator('.dock-container .rh-br').first();
      const box = await handle.boundingBox();

      await page.mouse.move(box.x + 2, box.y + 2);
      await page.mouse.down({ button: 'left' });
      await page.mouse.move(box.x + 52, box.y + 52);
      await page.mouse.up({ button: 'left' });

      const newSize = await page.evaluate(() => {
        const dock = app._test.dockContainers[0];
        return { width: parseInt(dock.el.style.width), height: parseInt(dock.el.style.height) };
      });

      expect(newSize.width).toBeGreaterThan(origSize.width);
      expect(newSize.height).toBeGreaterThan(origSize.height);
    });
  });
});
