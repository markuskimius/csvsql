const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, executeSQL } = require('../helpers');

test.describe('Display transforms in docked panes', () => {

  async function loadPluginConfig(page, config) {
    await page.evaluate((cfg) => {
      const errors = app._test.validatePlugin(cfg);
      if (errors.length) throw new Error(errors.join(', '));
      const compiled = app._test.compilePlugin(cfg);
      cfg._compiled = compiled;
      app._test.plugins.push(cfg);
      app._test.rebuildTransformCache();
    }, config);
  }

  // Click a cell in a specific table using Playwright's mouse (full mousedown
  // sequence). Brings the target to front so the click isn't intercepted.
  async function clickWindowCell(page, tableName, colIdx = 0, rowIdx = 0) {
    const rect = await page.evaluate(({ name, col, row }) => {
      const w = app._test.windows.find(w => w.tableName === name);
      if (!w) return null;
      if (w.dockId) {
        const dock = app._test.dockContainers.find(d => d.id === w.dockId);
        if (dock) dock.el.style.zIndex = 999999;
      } else {
        w.el.style.zIndex = 999999;
      }
      const container = w.bodyEl;
      const rows = container.querySelectorAll('table tbody tr:not(.virtual-pad)');
      if (!rows[row]) return null;
      const cells = rows[row].querySelectorAll('td.data-cell');
      if (!cells[col]) return null;
      const r = cells[col].getBoundingClientRect();
      return { x: r.x + r.width / 2, y: r.y + r.height / 2 };
    }, { name: tableName, col: colIdx, row: rowIdx });
    if (rect) await page.mouse.click(rect.x, rect.y);
  }

  // Get a cell's text content from a specific table (works docked or standalone)
  async function getCellText(page, tableName, colIdx = 0, rowIdx = 0) {
    return page.evaluate(({ name, col, row }) => {
      const w = app._test.windows.find(w => w.tableName === name);
      if (!w) return null;
      const container = w.bodyEl;
      const rows = container.querySelectorAll('table tbody tr:not(.virtual-pad)');
      if (!rows[row]) return null;
      const cells = rows[row].querySelectorAll('td.data-cell');
      return cells[col] ? cells[col].textContent : null;
    }, { name: tableName, col: colIdx, row: rowIdx });
  }

  // Check if a cell has contenteditable="true"
  async function isCellEditable(page, tableName, colIdx = 0, rowIdx = 0) {
    return page.evaluate(({ name, col, row }) => {
      const w = app._test.windows.find(w => w.tableName === name);
      if (!w) return null;
      const container = w.bodyEl;
      const rows = container.querySelectorAll('table tbody tr:not(.virtual-pad)');
      if (!rows[row]) return null;
      const cells = rows[row].querySelectorAll('td.data-cell');
      return cells[col] ? cells[col].getAttribute('contenteditable') === 'true' : null;
    }, { name: tableName, col: colIdx, row: rowIdx });
  }

  async function setupSplitDockWithPlugin(page) {
    await openApp(page);
    await uploadFile(page, 'sample1.csv');
    await waitForWindow(page, 'sample1');
    await executeSQL(page, "SELECT name, email INTO sample2 FROM sample1");
    await waitForWindow(page, 'sample2');

    // Load a display transform plugin
    await loadPluginConfig(page, {
      name: 'Upper Names', table: '.*',
      columns: [{ match: '^name$', display: 'upper(value)' }]
    });

    // Merge into a split dock
    await page.evaluate(() => {
      const wins = app._test.windows;
      app._test.mergeWindowsAsSplit(wins[0].id, wins[1].id, 'right');
    });

    // Re-render to apply transforms
    await page.evaluate(() => app._test.rerenderAllWindows());
    await page.waitForTimeout(200);
  }

  // ===========================================================================
  // 1. Display transforms render in docked panes
  // ===========================================================================
  test.describe('Transform rendering in docked panes', () => {

    test('display transform is applied to cells in split dock panes', async ({ page }) => {
      await setupSplitDockWithPlugin(page);

      const text1 = await getCellText(page, 'sample1', 0, 0);
      const text2 = await getCellText(page, 'sample2', 0, 0);
      expect(text1).toBe('ALICE JOHNSON');
      expect(text2).toBe('ALICE JOHNSON');
    });
  });

  // ===========================================================================
  // 2. Edit mode restores raw value in docked panes
  // ===========================================================================
  test.describe('Edit mode in docked panes', () => {

    test('F2 enters edit mode and shows raw value', async ({ page }) => {
      await setupSplitDockWithPlugin(page);

      // Verify transform is active
      expect(await getCellText(page, 'sample1', 0, 0)).toBe('ALICE JOHNSON');

      // Click and enter edit mode
      await clickWindowCell(page, 'sample1', 0, 0);
      await page.keyboard.press('F2');
      await page.waitForTimeout(100);

      // Cell should show raw value in edit mode
      const rawText = await getCellText(page, 'sample1', 0, 0);
      expect(rawText).toBe('Alice Johnson');
      expect(await isCellEditable(page, 'sample1', 0, 0)).toBe(true);
    });

    test('F2 key enters edit mode and shows raw value', async ({ page }) => {
      await setupSplitDockWithPlugin(page);

      await clickWindowCell(page, 'sample1', 0, 0);
      await page.keyboard.press('F2');
      await page.waitForTimeout(100);

      expect(await getCellText(page, 'sample1', 0, 0)).toBe('Alice Johnson');
      expect(await isCellEditable(page, 'sample1', 0, 0)).toBe(true);
    });

    test('i key enters edit mode and shows raw value', async ({ page }) => {
      await setupSplitDockWithPlugin(page);

      await clickWindowCell(page, 'sample1', 0, 0);
      await page.keyboard.press('i');
      await page.waitForTimeout(100);

      expect(await getCellText(page, 'sample1', 0, 0)).toBe('Alice Johnson');
      expect(await isCellEditable(page, 'sample1', 0, 0)).toBe(true);
    });

    test('edit mode works in the second pane of a split dock', async ({ page }) => {
      await setupSplitDockWithPlugin(page);

      expect(await getCellText(page, 'sample2', 0, 0)).toBe('ALICE JOHNSON');

      await clickWindowCell(page, 'sample2', 0, 0);
      await page.keyboard.press('F2');
      await page.waitForTimeout(100);

      expect(await getCellText(page, 'sample2', 0, 0)).toBe('Alice Johnson');
    });
  });

  // ===========================================================================
  // 3. Escape reverts edit and re-applies transform in docked panes
  // ===========================================================================
  test.describe('Escape in docked panes', () => {

    test('Escape reverts edit and re-applies display transform', async ({ page }) => {
      await setupSplitDockWithPlugin(page);

      await clickWindowCell(page, 'sample1', 0, 0);
      await page.keyboard.press('F2');
      await page.waitForTimeout(100);
      expect(await getCellText(page, 'sample1', 0, 0)).toBe('Alice Johnson');

      // Type something then Escape to revert
      await page.keyboard.type('Modified');
      await page.keyboard.press('Escape');
      await page.waitForTimeout(100);

      // Should revert to display-transformed original value
      expect(await getCellText(page, 'sample1', 0, 0)).toBe('ALICE JOHNSON');
      expect(await isCellEditable(page, 'sample1', 0, 0)).toBe(false);
    });
  });

  // ===========================================================================
  // 4. Blur (clicking another cell) re-applies transform in docked panes
  // ===========================================================================
  test.describe('Blur in docked panes', () => {

    test('clicking another cell after editing re-applies display transform', async ({ page }) => {
      await setupSplitDockWithPlugin(page);

      await clickWindowCell(page, 'sample1', 0, 0);
      await page.keyboard.press('F2');
      await page.waitForTimeout(100);
      expect(await getCellText(page, 'sample1', 0, 0)).toBe('Alice Johnson');

      // Click a different cell to blur
      await clickWindowCell(page, 'sample1', 1, 0);
      await page.waitForTimeout(100);

      // Original cell should be back to display-transformed value
      expect(await getCellText(page, 'sample1', 0, 0)).toBe('ALICE JOHNSON');
    });

    test('clicking a cell in the other pane re-applies display transform', async ({ page }) => {
      await setupSplitDockWithPlugin(page);

      await clickWindowCell(page, 'sample1', 0, 0);
      await page.keyboard.press('F2');
      await page.waitForTimeout(100);
      expect(await getCellText(page, 'sample1', 0, 0)).toBe('Alice Johnson');

      // Click a cell in the OTHER pane
      await clickWindowCell(page, 'sample2', 0, 0);
      await page.waitForTimeout(100);

      // Original cell in sample1 should be back to display-transformed value
      expect(await getCellText(page, 'sample1', 0, 0)).toBe('ALICE JOHNSON');
    });

    test('editing then blurring saves the value and applies transform', async ({ page }) => {
      await setupSplitDockWithPlugin(page);

      await clickWindowCell(page, 'sample1', 0, 0);
      await page.keyboard.press('F2');
      await page.waitForTimeout(100);

      // Clear cell and type a new value
      await page.evaluate(({ name }) => {
        const w = app._test.windows.find(w => w.tableName === name);
        const cell = w.bodyEl.querySelector('td.data-cell[contenteditable="true"]');
        if (cell) cell.textContent = '';
      }, { name: 'sample1' });
      await page.keyboard.type('jane doe');
      await clickWindowCell(page, 'sample1', 1, 0);
      await page.waitForTimeout(100);

      // Should show the new value display-transformed (uppercased)
      expect(await getCellText(page, 'sample1', 0, 0)).toBe('JANE DOE');

      // Underlying data should have the raw value
      const rawValue = await page.evaluate(() => {
        return app._test.tables['sample1'].rows[0].name;
      });
      expect(rawValue).toBe('jane doe');
    });
  });

  // ===========================================================================
  // 5. Tab dock (tabbed windows) also restores raw value on edit
  // ===========================================================================
  test.describe('Tab dock with transforms', () => {

    test('edit mode shows raw value in a tabbed dock pane', async ({ page }) => {
      await openApp(page);
      await uploadFile(page, 'sample1.csv');
      await waitForWindow(page, 'sample1');
      await executeSQL(page, "SELECT name, email INTO sample2 FROM sample1");
      await waitForWindow(page, 'sample2');

      await loadPluginConfig(page, {
        name: 'Upper Names', table: '.*',
        columns: [{ match: '^name$', display: 'upper(value)' }]
      });

      // Tab the windows together
      await page.evaluate(() => {
        const wins = app._test.windows;
        app._test.mergeWindowsIntoTabs(wins[0].id, wins[1].id);
      });

      await page.evaluate(() => app._test.rerenderAllWindows());
      await page.waitForTimeout(200);

      // Verify transform is showing
      expect(await getCellText(page, 'sample1', 0, 0)).toBe('ALICE JOHNSON');

      // Enter edit mode
      await clickWindowCell(page, 'sample1', 0, 0);
      await page.keyboard.press('F2');
      await page.waitForTimeout(100);

      // Should show raw value
      expect(await getCellText(page, 'sample1', 0, 0)).toBe('Alice Johnson');

      // Escape to revert
      await page.keyboard.press('Escape');
      await page.waitForTimeout(100);
      expect(await getCellText(page, 'sample1', 0, 0)).toBe('ALICE JOHNSON');
    });
  });

  // ===========================================================================
  // 6. Three-way split dock with transforms
  // ===========================================================================
  test.describe('Three-way split with transforms', () => {

    test('edit mode restores raw value in each pane of a 3-way split', async ({ page }) => {
      await setupSplitDockWithPlugin(page);

      // Add a third table
      await executeSQL(page, "SELECT name, email INTO sample3 FROM sample1");
      await waitForWindow(page, 'sample3');

      await page.evaluate(() => {
        const wins = app._test.windows;
        const dock = app._test.dockContainers[0];
        const w2 = wins.find(w => w.tableName === 'sample2');
        const w3 = wins.find(w => w.tableName === 'sample3');
        app._test.dockWindowAsSplit(w3.id, w2.dockLeaf, dock, 'bottom');
      });
      await page.evaluate(() => app._test.rerenderAllWindows());
      await page.waitForTimeout(200);

      // Test edit mode in all three panes
      for (const name of ['sample1', 'sample2', 'sample3']) {
        expect(await getCellText(page, name, 0, 0)).toBe('ALICE JOHNSON');

        await clickWindowCell(page, name, 0, 0);
        await page.keyboard.press('F2');
        await page.waitForTimeout(100);
        expect(await getCellText(page, name, 0, 0)).toBe('Alice Johnson');

        await page.keyboard.press('Escape');
        await page.waitForTimeout(100);
        expect(await getCellText(page, name, 0, 0)).toBe('ALICE JOHNSON');
      }
    });
  });

  // ===========================================================================
  // 7. Standalone window still works (regression check)
  // ===========================================================================
  test.describe('Standalone window regression', () => {

    test('edit mode still shows raw value in standalone window', async ({ page }) => {
      await openApp(page);
      await uploadFile(page, 'sample1.csv');
      await waitForWindow(page, 'sample1');

      await loadPluginConfig(page, {
        name: 'Upper Names', table: '.*',
        columns: [{ match: '^name$', display: 'upper(value)' }]
      });
      await page.evaluate(() => app._test.rerenderAllWindows());
      await page.waitForTimeout(200);

      expect(await getCellText(page, 'sample1', 0, 0)).toBe('ALICE JOHNSON');

      await clickWindowCell(page, 'sample1', 0, 0);
      await page.keyboard.press('F2');
      await page.waitForTimeout(100);
      expect(await getCellText(page, 'sample1', 0, 0)).toBe('Alice Johnson');

      await page.keyboard.press('Escape');
      await page.waitForTimeout(100);
      expect(await getCellText(page, 'sample1', 0, 0)).toBe('ALICE JOHNSON');
    });
  });
});
