const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, executeSQL } = require('../helpers');

test.describe('Dock / Tab System', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, 'sample1.csv');
    await waitForWindow(page, 'sample1');
    await executeSQL(page, "SELECT * INTO [query_result] FROM [sample1]");
    await waitForWindow(page, 'query_result');
  });

  test('mergeWindowsIntoTabs creates a dock container with tabs', async ({ page }) => {
    const result = await page.evaluate(() => {
      const wins = app._test.windows;
      const id0 = wins[0].id, id1 = wins[1].id;
      app._test.mergeWindowsIntoTabs(id0, id1);
      return {
        dockCount: app._test.dockContainers.length,
        win0docked: wins[0].dockId !== null,
        win1docked: wins[1].dockId !== null,
        tabCount: document.querySelectorAll('.dock-tab').length,
        containerVisible: document.querySelector('.dock-container') !== null,
        win0hidden: wins[0].el.classList.contains('docked'),
        win1hidden: wins[1].el.classList.contains('docked'),
      };
    });
    expect(result.dockCount).toBe(1);
    expect(result.win0docked).toBe(true);
    expect(result.win1docked).toBe(true);
    expect(result.tabCount).toBe(2);
    expect(result.containerVisible).toBe(true);
    expect(result.win0hidden).toBe(true);
    expect(result.win1hidden).toBe(true);
  });

  test('clicking tabs switches content', async ({ page }) => {
    await page.evaluate(() => {
      const wins = app._test.windows;
      app._test.mergeWindowsIntoTabs(wins[0].id, wins[1].id);
    });
    await page.waitForTimeout(100);

    // Click the first tab
    await page.click('.dock-tab:first-of-type');
    await page.waitForTimeout(100);

    const activeTitle = await page.evaluate(() => {
      const activeTab = document.querySelector('.dock-tab.active');
      return activeTab ? activeTab.querySelector('.dock-tab-title').textContent : null;
    });
    expect(activeTitle).toBeTruthy();
  });

  test('closing a tab removes it from dock', async ({ page }) => {
    page.on('dialog', dialog => dialog.accept());

    await page.evaluate(() => {
      const wins = app._test.windows;
      app._test.mergeWindowsIntoTabs(wins[0].id, wins[1].id);
    });
    await page.waitForTimeout(100);

    // Close one tab via the close button
    await page.click('.dock-tab .dock-tab-close');
    await page.waitForTimeout(200);

    // Dock should dissolve back to a standalone window
    const result = await page.evaluate(() => ({
      dockCount: app._test.dockContainers.length,
      windowCount: app._test.windows.length,
      standaloneCount: app._test.windows.filter(w => !w.dockId).length,
    }));
    expect(result.dockCount).toBe(0);
    expect(result.windowCount).toBe(1);
    expect(result.standaloneCount).toBe(1);
  });

  test('mergeWindowsAsSplit creates a split dock', async ({ page }) => {
    const result = await page.evaluate(() => {
      const wins = app._test.windows;
      app._test.mergeWindowsAsSplit(wins[0].id, wins[1].id, 'right');
      return {
        dockCount: app._test.dockContainers.length,
        splitExists: document.querySelector('.dock-split') !== null,
        splitterExists: document.querySelector('.dock-splitter') !== null,
        leafCount: document.querySelectorAll('.dock-leaf').length,
      };
    });
    expect(result.dockCount).toBe(1);
    expect(result.splitExists).toBe(true);
    expect(result.splitterExists).toBe(true);
    expect(result.leafCount).toBe(2);
  });

  test('undockWindow restores standalone window', async ({ page }) => {
    await page.evaluate(() => {
      const wins = app._test.windows;
      app._test.mergeWindowsIntoTabs(wins[0].id, wins[1].id);
    });

    const result = await page.evaluate(() => {
      const wins = app._test.windows;
      const undockId = wins[0].id;
      app._test.undockWindow(undockId);
      const undockedWin = wins.find(w => w.id === undockId);
      return {
        dockCount: app._test.dockContainers.length,
        undockedDockId: undockedWin ? undockedWin.dockId : 'not found',
        standaloneCount: wins.filter(w => !w.dockId).length,
      };
    });
    // After undocking one of two tabs, dock dissolves (single tab left)
    expect(result.dockCount).toBe(0);
    expect(result.undockedDockId).toBeNull();
    expect(result.standaloneCount).toBe(2);
  });
});
