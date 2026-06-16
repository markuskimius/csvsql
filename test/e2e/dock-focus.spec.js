const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, executeSQL, getTableData } = require('../helpers');

test.describe('Dock Focus Tracking', () => {

  // Helper: click a cell in a specific table using Playwright's mouse to trigger
  // the full mousedown -> mouseup -> click sequence. Works for both docked and
  // standalone windows. Brings standalone windows to front first so the click
  // isn't intercepted by an overlapping dock container.
  async function clickWindowCell(page, tableName) {
    const rect = await page.evaluate((name) => {
      const w = app._test.windows.find(w => w.tableName === name);
      if (!w) return null;
      if (w.dockId) {
        const dock = app._test.dockContainers.find(d => d.id === w.dockId);
        if (dock) dock.el.style.zIndex = 999999;
      } else {
        w.el.style.zIndex = 999999;
      }
      const container = w.dockLeaf ? w.dockLeaf.contentEl : w.bodyEl;
      const cell = container.querySelector('td.data-cell');
      if (!cell) return null;
      const r = cell.getBoundingClientRect();
      return { x: r.x + r.width / 2, y: r.y + r.height / 2 };
    }, tableName);
    if (rect) await page.mouse.click(rect.x, rect.y);
  }

  async function setupSplitDock(page) {
    await openApp(page);
    await uploadFile(page, 'sample1.csv');
    await waitForWindow(page, 'sample1');
    await executeSQL(page, "SELECT name, email INTO sample2 FROM sample1");
    await waitForWindow(page, 'sample2');
    await executeSQL(page, "SELECT name, email INTO sample3 FROM sample1");
    await waitForWindow(page, 'sample3');

    await page.evaluate(() => {
      const wins = app._test.windows;
      app._test.mergeWindowsAsSplit(wins[0].id, wins[1].id, 'right');
    });

    const dockCount = await page.evaluate(() => app._test.dockContainers.length);
    expect(dockCount).toBe(1);
  }

  async function setupTabDock(page) {
    await openApp(page);
    await uploadFile(page, 'sample1.csv');
    await waitForWindow(page, 'sample1');
    await executeSQL(page, "SELECT name, email INTO sample2 FROM sample1");
    await waitForWindow(page, 'sample2');

    await page.evaluate(() => {
      const wins = app._test.windows;
      app._test.mergeWindowsIntoTabs(wins[0].id, wins[1].id);
    });

    const dockCount = await page.evaluate(() => app._test.dockContainers.length);
    expect(dockCount).toBe(1);
  }

  // ===========================================================================
  // 1. activeWinId tracks correctly in split docks
  // ===========================================================================
  test.describe('Split dock focus', () => {

    test('clicking a cell in pane A sets activeWinId to that pane', async ({ page }) => {
      await setupSplitDock(page);
      await clickWindowCell(page, 'sample1');

      const match = await page.evaluate(() => {
        const w = app._test.windows.find(w => w.tableName === 'sample1');
        return app._test.activeWinId === w.id;
      });
      expect(match).toBe(true);
    });

    test('clicking a cell in pane B sets activeWinId to that pane', async ({ page }) => {
      await setupSplitDock(page);
      await clickWindowCell(page, 'sample2');

      const match = await page.evaluate(() => {
        const w = app._test.windows.find(w => w.tableName === 'sample2');
        return app._test.activeWinId === w.id;
      });
      expect(match).toBe(true);
    });

    test('clicking between panes alternates activeWinId correctly', async ({ page }) => {
      await setupSplitDock(page);

      for (const name of ['sample2', 'sample1', 'sample2']) {
        await clickWindowCell(page, name);
        const match = await page.evaluate((expected) => {
          const w = app._test.windows.find(w => w.tableName === expected);
          return app._test.activeWinId === w.id;
        }, name);
        expect(match).toBe(true);
      }
    });

    test('clicking the splitter sets a valid activeWinId via fallback', async ({ page }) => {
      await setupSplitDock(page);

      const splitter = page.locator('.dock-splitter').first();
      await splitter.click();

      const result = await page.evaluate(() => {
        const activeWin = app._test.windows.find(w => w.id === app._test.activeWinId);
        return {
          isValid: activeWin !== undefined,
          isDocked: activeWin ? activeWin.dockId !== null : false,
        };
      });
      expect(result.isValid).toBe(true);
      expect(result.isDocked).toBe(true);
    });
  });

  // ===========================================================================
  // 2. getActiveDataWindow returns the correct window in split docks
  // ===========================================================================
  test.describe('getActiveDataWindow in split docks', () => {

    test('returns pane A after clicking pane A', async ({ page }) => {
      await setupSplitDock(page);
      await clickWindowCell(page, 'sample1');

      const name = await page.evaluate(() => app._test.getActiveDataWindow()?.tableName);
      expect(name).toBe('sample1');
    });

    test('returns pane B after clicking pane B', async ({ page }) => {
      await setupSplitDock(page);
      await clickWindowCell(page, 'sample2');

      const name = await page.evaluate(() => app._test.getActiveDataWindow()?.tableName);
      expect(name).toBe('sample2');
    });

    test('switches correctly when clicking back and forth', async ({ page }) => {
      await setupSplitDock(page);

      for (const expected of ['sample1', 'sample2', 'sample1']) {
        await clickWindowCell(page, expected);
        const name = await page.evaluate(() => app._test.getActiveDataWindow()?.tableName);
        expect(name).toBe(expected);
      }
    });
  });

  // ===========================================================================
  // 3. Clipboard operations target the correct pane in split docks
  // ===========================================================================
  test.describe('Clipboard in split docks', () => {

    test.beforeEach(async ({ context }) => {
      await context.grantPermissions(['clipboard-read', 'clipboard-write']);
    });

    test('Ctrl+C copies from the clicked pane, not the other', async ({ page }) => {
      await setupSplitDock(page);

      await page.evaluate(() => {
        ['sample1', 'sample2'].forEach(name => {
          const w = app._test.windows.find(w => w.tableName === name);
          app._test.tables[name].rows[0].name = name === 'sample1' ? 'LEFT_VALUE' : 'RIGHT_VALUE';
          app._test.rebuildTable(w);
        });
      });

      await clickWindowCell(page, 'sample2');
      await page.keyboard.press('Control+c');
      let clip = await page.evaluate(() => navigator.clipboard.readText());
      expect(clip).toBe('RIGHT_VALUE');

      await clickWindowCell(page, 'sample1');
      await page.keyboard.press('Control+c');
      clip = await page.evaluate(() => navigator.clipboard.readText());
      expect(clip).toBe('LEFT_VALUE');
    });

    test('Ctrl+V pastes into the clicked pane only', async ({ page }) => {
      await setupSplitDock(page);

      await clickWindowCell(page, 'sample2');
      await page.evaluate(() => navigator.clipboard.writeText('PASTED_RIGHT'));
      await page.keyboard.press('Control+v');
      await page.waitForTimeout(300);

      const data = await page.evaluate(() => ({
        sample2: app._test.tables['sample2'].rows[0].name,
        sample1: app._test.tables['sample1'].rows[0].name,
      }));
      expect(data.sample2).toBe('PASTED_RIGHT');
      expect(data.sample1).not.toBe('PASTED_RIGHT');
    });

    test('Ctrl+V pastes into sample1 pane, not sample2', async ({ page }) => {
      await setupSplitDock(page);

      await clickWindowCell(page, 'sample1');
      await page.evaluate(() => navigator.clipboard.writeText('PASTED_LEFT'));
      await page.keyboard.press('Control+v');
      await page.waitForTimeout(300);

      const data = await page.evaluate(() => ({
        sample1: app._test.tables['sample1'].rows[0].name,
        sample2: app._test.tables['sample2'].rows[0].name,
      }));
      expect(data.sample1).toBe('PASTED_LEFT');
      expect(data.sample2).not.toBe('PASTED_LEFT');
    });

    test('Ctrl+X cuts from the clicked pane only', async ({ page }) => {
      await setupSplitDock(page);

      await page.evaluate(() => {
        ['sample1', 'sample2'].forEach(name => {
          const w = app._test.windows.find(w => w.tableName === name);
          app._test.tables[name].rows[0].name = 'CUT_' + name.toUpperCase();
          app._test.rebuildTable(w);
        });
      });

      await clickWindowCell(page, 'sample2');
      await page.keyboard.press('Control+x');
      await page.waitForTimeout(300);

      const data = await page.evaluate(() => ({
        sample2: app._test.tables['sample2'].rows[0].name,
        sample1: app._test.tables['sample1'].rows[0].name,
      }));
      expect(data.sample2).toBe('');
      expect(data.sample1).toBe('CUT_SAMPLE1');
    });
  });

  // ===========================================================================
  // 4. Tab dock focus (single-leaf dock with multiple tabs)
  // ===========================================================================
  test.describe('Tab dock focus', () => {

    test('clicking a tab switches activeWinId', async ({ page }) => {
      await setupTabDock(page);

      const tabInfo = await page.evaluate(() => {
        const dock = app._test.dockContainers[0];
        const leaf = dock.root;
        return leaf.tabs.map(id => {
          const w = app._test.windows.find(w => w.id === id);
          return { id, tableName: w.tableName };
        });
      });

      const sample2Idx = tabInfo.findIndex(t => t.tableName === 'sample2');
      const sample1Idx = tabInfo.findIndex(t => t.tableName === 'sample1');

      await page.locator('.dock-tab').nth(sample2Idx).click();
      let match = await page.evaluate(() => {
        const w = app._test.windows.find(w => w.tableName === 'sample2');
        return app._test.activeWinId === w.id;
      });
      expect(match).toBe(true);

      await page.locator('.dock-tab').nth(sample1Idx).click();
      match = await page.evaluate(() => {
        const w = app._test.windows.find(w => w.tableName === 'sample1');
        return app._test.activeWinId === w.id;
      });
      expect(match).toBe(true);
    });

    test('clicking a tab in dock updates activeWinId after standalone had focus', async ({ page }) => {
      await setupTabDock(page);

      await executeSQL(page, "SELECT name INTO sample3 FROM sample1");
      await waitForWindow(page, 'sample3');

      await clickWindowCell(page, 'sample3');
      let activeName = await page.evaluate(() => app._test.getActiveDataWindow()?.tableName);
      expect(activeName).toBe('sample3');

      // Click a dock tab to switch focus back to the dock
      const tabInfo = await page.evaluate(() => {
        const dock = app._test.dockContainers[0];
        const leaf = dock.root;
        return leaf.tabs.map(id => {
          const w = app._test.windows.find(w => w.id === id);
          return w.tableName;
        });
      });

      // Bring dock to front and click its first tab
      await page.evaluate(() => {
        const dock = app._test.dockContainers[0];
        dock.el.style.zIndex = 999999;
      });
      await page.locator('.dock-tab').first().click();
      activeName = await page.evaluate(() => app._test.getActiveDataWindow()?.tableName);
      expect(activeName).toBe(tabInfo[0]);
    });
  });

  // ===========================================================================
  // 5. Focus tracking with 3-way split docks
  // ===========================================================================
  test.describe('Three-way split dock focus', () => {

    async function setup3WaySplit(page) {
      await setupSplitDock(page);
      await page.evaluate(() => {
        const wins = app._test.windows;
        const dock = app._test.dockContainers[0];
        const w2 = wins.find(w => w.tableName === 'sample2');
        const w3 = wins.find(w => w.tableName === 'sample3');
        app._test.dockWindowAsSplit(w3.id, w2.dockLeaf, dock, 'bottom');
      });
    }

    test('clicking any of three panes sets activeWinId correctly', async ({ page }) => {
      await setup3WaySplit(page);

      const leafCount = await page.evaluate(() => {
        function count(node) {
          if (node.type === 'leaf') return 1;
          return node.children.reduce((s, c) => s + count(c), 0);
        }
        return count(app._test.dockContainers[0].root);
      });
      expect(leafCount).toBe(3);

      for (const name of ['sample1', 'sample2', 'sample3']) {
        await clickWindowCell(page, name);
        const activeName = await page.evaluate(() => app._test.getActiveDataWindow()?.tableName);
        expect(activeName).toBe(name);
      }
    });

    test('each pane in a 3-way split tracks independently for copy', async ({ page, context }) => {
      await context.grantPermissions(['clipboard-read', 'clipboard-write']);
      await setup3WaySplit(page);

      await page.evaluate(() => {
        ['sample1', 'sample2', 'sample3'].forEach(name => {
          const w = app._test.windows.find(w => w.tableName === name);
          app._test.tables[name].rows[0].name = 'PANE_' + name;
          app._test.rebuildTable(w);
        });
      });

      for (const name of ['sample1', 'sample2', 'sample3']) {
        await clickWindowCell(page, name);
        await page.keyboard.press('Control+c');
        const clip = await page.evaluate(() => navigator.clipboard.readText());
        expect(clip).toBe('PANE_' + name);
      }
    });

    test('copy after rapid pane switching lands on the last-clicked pane', async ({ page, context }) => {
      await context.grantPermissions(['clipboard-read', 'clipboard-write']);
      await setup3WaySplit(page);

      await page.evaluate(() => {
        ['sample1', 'sample2', 'sample3'].forEach(name => {
          const w = app._test.windows.find(w => w.tableName === name);
          app._test.tables[name].rows[0].name = 'VAL_' + name;
          app._test.rebuildTable(w);
        });
      });

      await clickWindowCell(page, 'sample1');
      await clickWindowCell(page, 'sample2');
      await clickWindowCell(page, 'sample3');
      await page.keyboard.press('Control+c');
      const clip = await page.evaluate(() => navigator.clipboard.readText());
      expect(clip).toBe('VAL_sample3');
    });
  });

  // ===========================================================================
  // 6. Focus survives dock operations
  // ===========================================================================
  test.describe('Focus persistence', () => {

    test('activeWinId remains valid after adding a tab to a split dock', async ({ page }) => {
      await setupSplitDock(page);

      await clickWindowCell(page, 'sample1');
      const beforeName = await page.evaluate(() => app._test.getActiveDataWindow()?.tableName);
      expect(beforeName).toBe('sample1');

      await page.evaluate(() => {
        const wins = app._test.windows;
        const dock = app._test.dockContainers[0];
        const w2 = wins.find(w => w.tableName === 'sample2');
        const w3 = wins.find(w => w.tableName === 'sample3');
        app._test.dockWindowAsTab(w3.id, w2.dockLeaf, dock);
      });

      const isValid = await page.evaluate(() => {
        return app._test.windows.some(w => w.id === app._test.activeWinId);
      });
      expect(isValid).toBe(true);
    });

    test('focus switches to standalone window when clicked after dock', async ({ page }) => {
      await setupSplitDock(page);

      await clickWindowCell(page, 'sample2');
      let name = await page.evaluate(() => app._test.getActiveDataWindow()?.tableName);
      expect(name).toBe('sample2');

      await clickWindowCell(page, 'sample3');
      name = await page.evaluate(() => app._test.getActiveDataWindow()?.tableName);
      expect(name).toBe('sample3');
    });

    test('focus returns to dock pane after clicking standalone then dock', async ({ page }) => {
      await setupSplitDock(page);

      await clickWindowCell(page, 'sample3');
      let name = await page.evaluate(() => app._test.getActiveDataWindow()?.tableName);
      expect(name).toBe('sample3');

      await clickWindowCell(page, 'sample1');
      name = await page.evaluate(() => app._test.getActiveDataWindow()?.tableName);
      expect(name).toBe('sample1');
    });
  });

  // ===========================================================================
  // 7. Edit menu uses getActiveDataWindow correctly
  // ===========================================================================
  test.describe('Edit menu targets correct pane', () => {

    test.beforeEach(async ({ context }) => {
      await context.grantPermissions(['clipboard-read', 'clipboard-write']);
    });

    test('activeWinId matches clicked pane for both panes', async ({ page }) => {
      await setupSplitDock(page);

      await clickWindowCell(page, 'sample1');
      let match = await page.evaluate(() => {
        const w = app._test.windows.find(w => w.tableName === 'sample1');
        return app._test.activeWinId === w.id;
      });
      expect(match).toBe(true);

      await clickWindowCell(page, 'sample2');
      match = await page.evaluate(() => {
        const w = app._test.windows.find(w => w.tableName === 'sample2');
        return app._test.activeWinId === w.id;
      });
      expect(match).toBe(true);
    });

    test('getActiveDataWindow matches the pane where Ctrl+C operates', async ({ page }) => {
      await setupSplitDock(page);

      await page.evaluate(() => {
        ['sample1', 'sample2'].forEach(name => {
          const w = app._test.windows.find(w => w.tableName === name);
          app._test.tables[name].rows[0].name = 'COPY_' + name;
          app._test.rebuildTable(w);
        });
      });

      await clickWindowCell(page, 'sample1');
      await page.keyboard.press('Control+c');
      let clip = await page.evaluate(() => navigator.clipboard.readText());
      let activeName = await page.evaluate(() => app._test.getActiveDataWindow()?.tableName);
      expect(activeName).toBe('sample1');
      expect(clip).toBe('COPY_sample1');

      await clickWindowCell(page, 'sample2');
      await page.keyboard.press('Control+c');
      clip = await page.evaluate(() => navigator.clipboard.readText());
      activeName = await page.evaluate(() => app._test.getActiveDataWindow()?.tableName);
      expect(activeName).toBe('sample2');
      expect(clip).toBe('COPY_sample2');
    });
  });
});
