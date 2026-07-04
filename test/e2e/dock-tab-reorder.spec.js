const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, executeSQL } = require('../helpers');

test.describe('Dock Tab Reorder', () => {

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
  // 1. Basic reorder — drag tab to a new position
  // =========================================================================
  test.describe('Basic reorder', () => {

    test('dragging a tab left reorders it before the target', async ({ page }) => {
      const before = await page.evaluate(() =>
        app._test.dockContainers[0].root.tabs.slice()
      );

      // Drag the last tab to the left of the first tab
      const tabs = page.locator('.dock-tab');
      const lastTab = tabs.nth(2);
      const firstTab = tabs.nth(0);
      const lastBox = await lastTab.boundingBox();
      const firstBox = await firstTab.boundingBox();

      await page.mouse.move(lastBox.x + lastBox.width / 2, lastBox.y + lastBox.height / 2);
      await page.mouse.down({ button: 'left' });
      // Move past threshold
      await page.mouse.move(lastBox.x + lastBox.width / 2 - 10, lastBox.y + lastBox.height / 2, { steps: 3 });
      // Move to left side of first tab
      await page.mouse.move(firstBox.x + 5, firstBox.y + firstBox.height / 2, { steps: 5 });
      await page.mouse.up({ button: 'left' });

      const after = await page.evaluate(() =>
        app._test.dockContainers[0].root.tabs.slice()
      );

      // The last tab should now be first
      expect(after[0]).toBe(before[2]);
      expect(after[1]).toBe(before[0]);
      expect(after[2]).toBe(before[1]);
    });

    test('dragging a tab right reorders it after the target', async ({ page }) => {
      const before = await page.evaluate(() =>
        app._test.dockContainers[0].root.tabs.slice()
      );

      // Drag the first tab to the right of the last tab
      const tabs = page.locator('.dock-tab');
      const firstTab = tabs.nth(0);
      const lastTab = tabs.nth(2);
      const firstBox = await firstTab.boundingBox();
      const lastBox = await lastTab.boundingBox();

      await page.mouse.move(firstBox.x + firstBox.width / 2, firstBox.y + firstBox.height / 2);
      await page.mouse.down({ button: 'left' });
      await page.mouse.move(firstBox.x + firstBox.width / 2 + 10, firstBox.y + firstBox.height / 2, { steps: 3 });
      // Move to right side of last tab
      await page.mouse.move(lastBox.x + lastBox.width - 5, lastBox.y + lastBox.height / 2, { steps: 5 });
      await page.mouse.up({ button: 'left' });

      const after = await page.evaluate(() =>
        app._test.dockContainers[0].root.tabs.slice()
      );

      // First tab should now be last
      expect(after[0]).toBe(before[1]);
      expect(after[1]).toBe(before[2]);
      expect(after[2]).toBe(before[0]);
    });

    test('dragging to same position is a no-op', async ({ page }) => {
      const before = await page.evaluate(() =>
        app._test.dockContainers[0].root.tabs.slice()
      );

      // Drag first tab slightly to the right but within its own bounds
      const tabs = page.locator('.dock-tab');
      const firstTab = tabs.nth(0);
      const firstBox = await firstTab.boundingBox();

      await page.mouse.move(firstBox.x + firstBox.width / 2, firstBox.y + firstBox.height / 2);
      await page.mouse.down({ button: 'left' });
      // Move past threshold but stay over same tab
      await page.mouse.move(firstBox.x + firstBox.width / 2 + 8, firstBox.y + firstBox.height / 2, { steps: 3 });
      await page.mouse.up({ button: 'left' });

      const after = await page.evaluate(() =>
        app._test.dockContainers[0].root.tabs.slice()
      );

      expect(after).toEqual(before);
    });

    test('dropping past the last tab moves to end', async ({ page }) => {
      const before = await page.evaluate(() =>
        app._test.dockContainers[0].root.tabs.slice()
      );

      // Drag first tab far to the right past all tabs
      const tabs = page.locator('.dock-tab');
      const firstTab = tabs.nth(0);
      const firstBox = await firstTab.boundingBox();
      const tabBar = page.locator('.dock-tab-bar').first();
      const barBox = await tabBar.boundingBox();

      await page.mouse.move(firstBox.x + firstBox.width / 2, firstBox.y + firstBox.height / 2);
      await page.mouse.down({ button: 'left' });
      await page.mouse.move(firstBox.x + firstBox.width / 2 + 10, firstBox.y + firstBox.height / 2, { steps: 3 });
      // Move far right past all tabs
      await page.mouse.move(barBox.x + barBox.width - 10, barBox.y + barBox.height / 2, { steps: 5 });
      await page.mouse.up({ button: 'left' });

      const after = await page.evaluate(() =>
        app._test.dockContainers[0].root.tabs.slice()
      );

      expect(after[0]).toBe(before[1]);
      expect(after[1]).toBe(before[2]);
      expect(after[2]).toBe(before[0]);
    });

    test('dropping before the first tab moves to start', async ({ page }) => {
      const before = await page.evaluate(() =>
        app._test.dockContainers[0].root.tabs.slice()
      );

      // Drag last tab far to the left
      const tabs = page.locator('.dock-tab');
      const lastTab = tabs.nth(2);
      const lastBox = await lastTab.boundingBox();
      const tabBar = page.locator('.dock-tab-bar').first();
      const barBox = await tabBar.boundingBox();

      await page.mouse.move(lastBox.x + lastBox.width / 2, lastBox.y + lastBox.height / 2);
      await page.mouse.down({ button: 'left' });
      await page.mouse.move(lastBox.x + lastBox.width / 2 - 10, lastBox.y + lastBox.height / 2, { steps: 3 });
      // Move far left before all tabs
      await page.mouse.move(barBox.x + 2, barBox.y + barBox.height / 2, { steps: 5 });
      await page.mouse.up({ button: 'left' });

      const after = await page.evaluate(() =>
        app._test.dockContainers[0].root.tabs.slice()
      );

      expect(after[0]).toBe(before[2]);
      expect(after[1]).toBe(before[0]);
      expect(after[2]).toBe(before[1]);
    });
  });

  // =========================================================================
  // 2. Visual feedback during reorder drag
  // =========================================================================
  test.describe('Visual feedback', () => {

    test('dragged tab gets dock-tab-dragging class', async ({ page }) => {
      const tabs = page.locator('.dock-tab');
      const firstTab = tabs.nth(0);
      const firstBox = await firstTab.boundingBox();

      await page.mouse.move(firstBox.x + firstBox.width / 2, firstBox.y + firstBox.height / 2);
      await page.mouse.down({ button: 'left' });
      // Move past threshold
      await page.mouse.move(firstBox.x + firstBox.width / 2 + 10, firstBox.y + firstBox.height / 2, { steps: 3 });

      const hasDragging = await firstTab.evaluate(el => el.classList.contains('dock-tab-dragging'));
      expect(hasDragging).toBe(true);

      await page.mouse.up({ button: 'left' });
    });

    test('dock-tab-dragging class is removed on mouseup', async ({ page }) => {
      const tabs = page.locator('.dock-tab');
      const firstTab = tabs.nth(0);
      const firstBox = await firstTab.boundingBox();

      await page.mouse.move(firstBox.x + firstBox.width / 2, firstBox.y + firstBox.height / 2);
      await page.mouse.down({ button: 'left' });
      await page.mouse.move(firstBox.x + firstBox.width / 2 + 10, firstBox.y + firstBox.height / 2, { steps: 3 });
      await page.mouse.up({ button: 'left' });

      const anyDragging = await page.evaluate(() =>
        document.querySelector('.dock-tab-dragging') !== null
      );
      expect(anyDragging).toBe(false);
    });

    test('drop indicator shows on target tab during drag', async ({ page }) => {
      const tabs = page.locator('.dock-tab');
      const firstTab = tabs.nth(0);
      const secondTab = tabs.nth(1);
      const firstBox = await firstTab.boundingBox();
      const secondBox = await secondTab.boundingBox();

      await page.mouse.move(firstBox.x + firstBox.width / 2, firstBox.y + firstBox.height / 2);
      await page.mouse.down({ button: 'left' });
      // Move past threshold
      await page.mouse.move(firstBox.x + firstBox.width / 2 + 10, firstBox.y + firstBox.height / 2, { steps: 3 });
      // Move over right side of second tab
      await page.mouse.move(secondBox.x + secondBox.width - 5, secondBox.y + secondBox.height / 2, { steps: 3 });

      const indicator = await page.evaluate(() => {
        const left = document.querySelector('.dock-tab-drop-left');
        const right = document.querySelector('.dock-tab-drop-right');
        return { hasLeft: left !== null, hasRight: right !== null };
      });

      expect(indicator.hasLeft || indicator.hasRight).toBe(true);

      await page.mouse.up({ button: 'left' });
    });

    test('drop indicators are cleared on mouseup', async ({ page }) => {
      const tabs = page.locator('.dock-tab');
      const firstTab = tabs.nth(0);
      const secondTab = tabs.nth(1);
      const firstBox = await firstTab.boundingBox();
      const secondBox = await secondTab.boundingBox();

      await page.mouse.move(firstBox.x + firstBox.width / 2, firstBox.y + firstBox.height / 2);
      await page.mouse.down({ button: 'left' });
      await page.mouse.move(firstBox.x + firstBox.width / 2 + 10, firstBox.y + firstBox.height / 2, { steps: 3 });
      await page.mouse.move(secondBox.x + secondBox.width / 2, secondBox.y + secondBox.height / 2, { steps: 3 });
      await page.mouse.up({ button: 'left' });

      const anyIndicators = await page.evaluate(() =>
        document.querySelector('.dock-tab-drop-left, .dock-tab-drop-right') !== null
      );
      expect(anyIndicators).toBe(false);
    });

    test('drop-left indicator shows when cursor is on left half of target', async ({ page }) => {
      const tabs = page.locator('.dock-tab');
      const lastTab = tabs.nth(2);
      const secondTab = tabs.nth(1);
      const lastBox = await lastTab.boundingBox();
      const secondBox = await secondTab.boundingBox();

      await page.mouse.move(lastBox.x + lastBox.width / 2, lastBox.y + lastBox.height / 2);
      await page.mouse.down({ button: 'left' });
      await page.mouse.move(lastBox.x + lastBox.width / 2 - 10, lastBox.y + lastBox.height / 2, { steps: 3 });
      // Move to left half of second tab
      await page.mouse.move(secondBox.x + 5, secondBox.y + secondBox.height / 2, { steps: 3 });

      const result = await page.evaluate(() => {
        const dropLeft = document.querySelector('.dock-tab-drop-left');
        return dropLeft !== null;
      });
      expect(result).toBe(true);

      await page.mouse.up({ button: 'left' });
    });

    test('drop-right indicator shows when cursor is on right half of target', async ({ page }) => {
      const tabs = page.locator('.dock-tab');
      const firstTab = tabs.nth(0);
      const secondTab = tabs.nth(1);
      const firstBox = await firstTab.boundingBox();
      const secondBox = await secondTab.boundingBox();

      await page.mouse.move(firstBox.x + firstBox.width / 2, firstBox.y + firstBox.height / 2);
      await page.mouse.down({ button: 'left' });
      await page.mouse.move(firstBox.x + firstBox.width / 2 + 10, firstBox.y + firstBox.height / 2, { steps: 3 });
      // Move to right half of second tab
      await page.mouse.move(secondBox.x + secondBox.width - 5, secondBox.y + secondBox.height / 2, { steps: 3 });

      const result = await page.evaluate(() => {
        const dropRight = document.querySelector('.dock-tab-drop-right');
        return dropRight !== null;
      });
      expect(result).toBe(true);

      await page.mouse.up({ button: 'left' });
    });
  });

  // =========================================================================
  // 3. Reorder does not interfere with other interactions
  // =========================================================================
  test.describe('Interaction isolation', () => {

    test('shift+drag still undocks instead of reordering', async ({ page }) => {
      const beforeTabs = await page.evaluate(() =>
        app._test.dockContainers[0].root.tabs.slice()
      );

      const tab = page.locator('.dock-tab').first();
      const tabBox = await tab.boundingBox();
      const dockEl = page.locator('.dock-container');
      const dockBox = await dockEl.boundingBox();

      // Shift-drag far outside the dock
      await page.mouse.move(tabBox.x + tabBox.width / 2, tabBox.y + tabBox.height / 2);
      await page.keyboard.down('Shift');
      await page.mouse.down({ button: 'left' });
      await page.mouse.move(dockBox.x + dockBox.width + 100, dockBox.y + dockBox.height + 100, { steps: 5 });
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

    test('reorder does not create ghost element', async ({ page }) => {
      const tabs = page.locator('.dock-tab');
      const firstTab = tabs.nth(0);
      const lastTab = tabs.nth(2);
      const firstBox = await firstTab.boundingBox();
      const lastBox = await lastTab.boundingBox();

      await page.mouse.move(firstBox.x + firstBox.width / 2, firstBox.y + firstBox.height / 2);
      await page.mouse.down({ button: 'left' });
      await page.mouse.move(lastBox.x + lastBox.width / 2, lastBox.y + lastBox.height / 2, { steps: 5 });

      const ghostExists = await page.evaluate(() =>
        document.querySelector('body > .subwindow[style*="position: fixed"]') !== null
      );
      expect(ghostExists).toBe(false);

      await page.mouse.up({ button: 'left' });
    });

    test('reorder does not move the dock container', async ({ page }) => {
      const origPos = await page.evaluate(() => {
        const dock = app._test.dockContainers[0];
        return { left: dock.el.style.left, top: dock.el.style.top };
      });

      const tabs = page.locator('.dock-tab');
      const firstTab = tabs.nth(0);
      const lastTab = tabs.nth(2);
      const firstBox = await firstTab.boundingBox();
      const lastBox = await lastTab.boundingBox();

      await page.mouse.move(firstBox.x + firstBox.width / 2, firstBox.y + firstBox.height / 2);
      await page.mouse.down({ button: 'left' });
      await page.mouse.move(lastBox.x + lastBox.width / 2, lastBox.y + lastBox.height / 2, { steps: 5 });
      await page.mouse.up({ button: 'left' });

      const newPos = await page.evaluate(() => {
        const dock = app._test.dockContainers[0];
        return { left: dock.el.style.left, top: dock.el.style.top };
      });

      expect(newPos.left).toBe(origPos.left);
      expect(newPos.top).toBe(origPos.top);
    });

    test('dragging empty tab bar area still moves dock', async ({ page }) => {
      const origPos = await page.evaluate(() => {
        const dock = app._test.dockContainers[0];
        return { left: parseInt(dock.el.style.left), top: parseInt(dock.el.style.top) };
      });

      // Find empty space between last tab and dock controls
      const grabPoint = await page.evaluate(() => {
        const tabBar = document.querySelector('.dock-tab-bar');
        const tabs = tabBar.querySelectorAll('.dock-tab');
        const lastTab = tabs[tabs.length - 1];
        const lastTabRect = lastTab.getBoundingClientRect();
        const controls = tabBar.querySelector('.dock-controls');
        const controlsRect = controls ? controls.getBoundingClientRect() : null;
        const rightEdge = controlsRect ? controlsRect.left : tabBar.getBoundingClientRect().right;
        // Midpoint between last tab and controls/bar edge
        const x = (lastTabRect.right + rightEdge) / 2;
        const y = tabBar.getBoundingClientRect().top + tabBar.getBoundingClientRect().height / 2;
        return { x, y };
      });

      await page.mouse.move(grabPoint.x, grabPoint.y);
      await page.mouse.down({ button: 'left' });
      await page.mouse.move(grabPoint.x + 50, grabPoint.y + 30, { steps: 5 });
      await page.mouse.up({ button: 'left' });

      const newPos = await page.evaluate(() => {
        const dock = app._test.dockContainers[0];
        return { left: parseInt(dock.el.style.left), top: parseInt(dock.el.style.top) };
      });

      expect(newPos.left).toBeGreaterThan(origPos.left);
      expect(newPos.top).toBeGreaterThan(origPos.top);
    });

    test('right-click does not trigger reorder', async ({ page }) => {
      const before = await page.evaluate(() =>
        app._test.dockContainers[0].root.tabs.slice()
      );

      const tabs = page.locator('.dock-tab');
      const firstTab = tabs.nth(0);
      const lastTab = tabs.nth(2);
      const firstBox = await firstTab.boundingBox();
      const lastBox = await lastTab.boundingBox();

      await page.mouse.move(firstBox.x + firstBox.width / 2, firstBox.y + firstBox.height / 2);
      await page.mouse.down({ button: 'right' });
      await page.mouse.move(lastBox.x + lastBox.width / 2, lastBox.y + lastBox.height / 2, { steps: 5 });
      await page.mouse.up({ button: 'right' });

      const after = await page.evaluate(() =>
        app._test.dockContainers[0].root.tabs.slice()
      );

      expect(after).toEqual(before);
    });
  });

  // =========================================================================
  // 4. Single-tab leaf behavior
  // =========================================================================
  test.describe('Single-tab leaf', () => {

    test('dragging a single tab never moves the dock — it undocks the tab', async ({ page }) => {
      // Split one of the tabs out of the existing 3-tab dock into a side split
      // This creates a dock with a split: one leaf has 2 tabs, the other has 1 tab
      const draggedId = await page.evaluate(() => {
        const dock = app._test.dockContainers[0];
        const leaf = dock.root;
        const winId = leaf.tabs[2];
        app._test.removeTabFromLeaf(winId, leaf, dock, true);
        app._test.dockWindowAsSplit(winId, dock.root, dock, 'right');
        return winId;
      });

      const orig = await page.evaluate(() => {
        const dock = app._test.dockContainers[0];
        const area = document.getElementById('window-area').getBoundingClientRect();
        const rect = dock.el.getBoundingClientRect();
        return {
          left: parseInt(dock.el.style.left), top: parseInt(dock.el.style.top),
          rect: { left: rect.left, top: rect.top, right: rect.right, bottom: rect.bottom },
          area: { left: area.left, top: area.top, right: area.right, bottom: area.bottom },
        };
      });

      // Find the single-tab leaf's tab element and drag it onto empty space
      const tab = page.locator('.dock-tab-bar.single-tab .dock-tab');
      const tabBox = await tab.boundingBox();
      // A drop point inside the window area but outside the dock
      const dropX = orig.rect.right + 60 < orig.area.right ? orig.rect.right + 60 : orig.rect.left - 60;
      const dropY = Math.min(orig.rect.bottom + 60, orig.area.bottom - 10);

      await page.mouse.move(tabBox.x + tabBox.width / 2, tabBox.y + tabBox.height / 2);
      await page.mouse.down({ button: 'left' });
      await page.mouse.move(dropX, dropY, { steps: 8 });
      await page.mouse.up({ button: 'left' });

      const after = await page.evaluate((winId) => {
        const dock = app._test.dockContainers[0];
        const win = app._test.windows.find(w => w.id === winId);
        return {
          left: parseInt(dock.el.style.left), top: parseInt(dock.el.style.top),
          draggedDockId: win.dockId,
        };
      }, draggedId);

      // The dock did not move; the tab undocked into a standalone window
      expect(after.left).toBe(orig.left);
      expect(after.top).toBe(orig.top);
      expect(after.draggedDockId).toBe(null);
    });
  });

  // =========================================================================
  // 5. Tab bar grab area
  // =========================================================================
  test.describe('Tab bar grab area', () => {

    test('tab bar has ::after pseudo-element with minimum width', async ({ page }) => {
      const afterWidth = await page.evaluate(() => {
        const tabBar = document.querySelector('.dock-tab-bar');
        const afterStyle = window.getComputedStyle(tabBar, '::after');
        return parseInt(afterStyle.minWidth);
      });

      expect(afterWidth).toBeGreaterThanOrEqual(40);
    });

    test('tab bar is wider than the sum of its tab children', async ({ page }) => {
      const result = await page.evaluate(() => {
        const tabBar = document.querySelector('.dock-tab-bar');
        const tabs = tabBar.querySelectorAll('.dock-tab');
        let tabWidthSum = 0;
        tabs.forEach(t => { tabWidthSum += t.offsetWidth; });
        return {
          barWidth: tabBar.scrollWidth,
          tabWidthSum,
        };
      });

      expect(result.barWidth).toBeGreaterThan(result.tabWidthSum);
    });
  });

  // =========================================================================
  // 6. Tab width behavior
  // =========================================================================
  test.describe('Tab width', () => {

    test('tabs are content-width, not stretched equally', async ({ page }) => {
      // The three windows have names: sample1, t2, t3 — different lengths
      // Tabs should have different widths since they fit their content
      const result = await page.evaluate(() => {
        const tabs = document.querySelectorAll('.dock-tab');
        const widths = [...tabs].map(t => t.offsetWidth);
        const allEqual = widths.every(w => w === widths[0]);
        return { widths, allEqual, count: tabs.length };
      });

      // With 3 tabs of different-length names, widths should not all be identical
      expect(result.count).toBe(3);
      expect(result.allEqual).toBe(false);
    });

    test('tabs use flex-shrink to fit narrow containers', async ({ page }) => {
      const result = await page.evaluate(() => {
        const tab = document.querySelector('.dock-tab');
        const style = window.getComputedStyle(tab);
        return {
          flexGrow: style.flexGrow,
          flexShrink: style.flexShrink,
        };
      });

      expect(result.flexGrow).toBe('0');
      expect(result.flexShrink).toBe('1');
    });
  });

  // =========================================================================
  // 7. Active tab preserved after reorder
  // =========================================================================
  test.describe('Active tab preservation', () => {

    test('dragged tab becomes active and stays active after reorder', async ({ page }) => {
      // The mousedown on a tab activates it, so after reorder the dragged tab should be active
      const draggedWinId = await page.evaluate(() =>
        app._test.dockContainers[0].root.tabs[2]
      );

      // Drag the last tab to the first position
      const tabs = page.locator('.dock-tab');
      const lastTab = tabs.nth(2);
      const firstTab = tabs.nth(0);
      const lastBox = await lastTab.boundingBox();
      const firstBox = await firstTab.boundingBox();

      await page.mouse.move(lastBox.x + lastBox.width / 2, lastBox.y + lastBox.height / 2);
      await page.mouse.down({ button: 'left' });
      await page.mouse.move(lastBox.x + lastBox.width / 2 - 10, lastBox.y + lastBox.height / 2, { steps: 3 });
      await page.mouse.move(firstBox.x + 5, firstBox.y + firstBox.height / 2, { steps: 5 });
      await page.mouse.up({ button: 'left' });

      const activeAfter = await page.evaluate(() =>
        app._test.dockContainers[0].root.activeTab
      );

      expect(activeAfter).toBe(draggedWinId);
    });

    test('active tab has .active class in DOM after reorder', async ({ page }) => {
      const activeWinId = await page.evaluate(() => {
        const dock = app._test.dockContainers[0];
        return dock.root.activeTab;
      });

      // Perform a drag reorder
      const tabs = page.locator('.dock-tab');
      const lastTab = tabs.nth(2);
      const firstTab = tabs.nth(0);
      const lastBox = await lastTab.boundingBox();
      const firstBox = await firstTab.boundingBox();

      await page.mouse.move(lastBox.x + lastBox.width / 2, lastBox.y + lastBox.height / 2);
      await page.mouse.down({ button: 'left' });
      await page.mouse.move(lastBox.x + lastBox.width / 2 - 10, lastBox.y + lastBox.height / 2, { steps: 3 });
      await page.mouse.move(firstBox.x + 5, firstBox.y + firstBox.height / 2, { steps: 5 });
      await page.mouse.up({ button: 'left' });

      const result = await page.evaluate((expectedId) => {
        const activeTabEl = document.querySelector('.dock-tab.active');
        return activeTabEl ? parseInt(activeTabEl.dataset.winId) : null;
      }, activeWinId);

      expect(result).toBe(activeWinId);
    });
  });
});
