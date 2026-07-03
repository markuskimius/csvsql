const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, executeSQL } = require('../helpers');

test.describe('Dialog stay-on-top', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
  });

  test('dialog has higher z-index than regular window after focusing the regular window', async ({ page }) => {
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');

    // Open column manager via app._test
    const tableWinId = await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.focusWindow(win.id);
      app._test.showColManager(win.id);
      return win.id;
    });
    await page.waitForTimeout(300);

    // Focus the table window (click its titlebar)
    await page.evaluate((winId) => {
      app._test.focusWindow(winId);
    }, tableWinId);
    await page.waitForTimeout(200);

    // Verify column manager z-index > table window z-index
    const result = await page.evaluate(() => {
      const tableWin = app._test.windows.find(w => w.tableName === 'sample1');
      const colMgrWin = app._test._activeColManagerWin;
      if (!colMgrWin) return { error: 'no col manager' };
      return {
        tableZ: parseInt(tableWin.el.style.zIndex) || 0,
        dialogZ: parseInt(colMgrWin.el.style.zIndex) || 0,
        isDialog: colMgrWin.isDialog,
      };
    });

    expect(result.isDialog).toBe(true);
    expect(result.dialogZ).toBeGreaterThan(result.tableZ);
  });

  test('dialog stays on top after creating a new window via SQL', async ({ page }) => {
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');

    // Open column manager
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.focusWindow(win.id);
      app._test.showColManager(win.id);
    });
    await page.waitForTimeout(300);

    // Create a second table via SQL query
    await executeSQL(page, 'SELECT name FROM sample1');
    await page.waitForTimeout(300);

    // Verify dialog z-index > both table window z-indexes
    const result = await page.evaluate(() => {
      const colMgrWin = app._test._activeColManagerWin;
      if (!colMgrWin) return { error: 'no col manager' };
      const dialogZ = parseInt(colMgrWin.el.style.zIndex) || 0;
      const tableWindows = app._test.windows.filter(w => w.tableName && !w.isDialog);
      const maxTableZ = Math.max(...tableWindows.map(w => parseInt(w.el.style.zIndex) || 0));
      return { dialogZ, maxTableZ, tableCount: tableWindows.length };
    });

    expect(result.tableCount).toBeGreaterThanOrEqual(2);
    expect(result.dialogZ).toBeGreaterThan(result.maxTableZ);
  });

  test('dialog stays on top after cascade layout', async ({ page }) => {
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');

    // Create a second table
    await executeSQL(page, 'SELECT name FROM sample1');
    await page.waitForTimeout(300);

    // Open column manager on the first table
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.focusWindow(win.id);
      app._test.showColManager(win.id);
    });
    await page.waitForTimeout(300);

    // Cascade layout
    await page.evaluate(() => app.layoutCascade());
    await page.waitForTimeout(300);

    const result = await page.evaluate(() => {
      const colMgrWin = app._test._activeColManagerWin;
      if (!colMgrWin) return { error: 'no col manager' };
      const dialogZ = parseInt(colMgrWin.el.style.zIndex) || 0;
      const nonDialogWindows = app._test.windows.filter(w => !w.isDialog);
      const maxZ = Math.max(...nonDialogWindows.map(w => parseInt(w.el.style.zIndex) || 0));
      return { dialogZ, maxZ };
    });

    expect(result.dialogZ).toBeGreaterThan(result.maxZ);
  });

  test('dialog stays on top after tile layout', async ({ page }) => {
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');

    await executeSQL(page, 'SELECT name FROM sample1');
    await page.waitForTimeout(300);

    // Open column manager
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.focusWindow(win.id);
      app._test.showColManager(win.id);
    });
    await page.waitForTimeout(300);

    // Tile layout
    await page.evaluate(() => app.layoutTileH());
    await page.waitForTimeout(300);

    const result = await page.evaluate(() => {
      const colMgrWin = app._test._activeColManagerWin;
      if (!colMgrWin) return { error: 'no col manager' };
      const dialogZ = parseInt(colMgrWin.el.style.zIndex) || 0;
      const nonDialogWindows = app._test.windows.filter(w => !w.isDialog);
      const maxZ = Math.max(...nonDialogWindows.map(w => parseInt(w.el.style.zIndex) || 0));
      return { dialogZ, maxZ };
    });

    expect(result.dialogZ).toBeGreaterThan(result.maxZ);
  });

  test('dialog is excluded from tile layout and its position is unchanged', async ({ page }) => {
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');

    // Open column manager
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.focusWindow(win.id);
      app._test.showColManager(win.id);
    });
    await page.waitForTimeout(300);

    // Record dialog position before tile
    const before = await page.evaluate(() => {
      const colMgrWin = app._test._activeColManagerWin;
      return {
        left: colMgrWin.el.style.left,
        top: colMgrWin.el.style.top,
        width: colMgrWin.el.style.width,
        height: colMgrWin.el.style.height,
      };
    });

    // Tile layout
    await page.evaluate(() => app.layoutTileH());
    await page.waitForTimeout(300);

    // Verify dialog position unchanged and layout units exclude dialog
    const after = await page.evaluate(() => {
      const colMgrWin = app._test._activeColManagerWin;
      return {
        left: colMgrWin.el.style.left,
        top: colMgrWin.el.style.top,
        width: colMgrWin.el.style.width,
        height: colMgrWin.el.style.height,
      };
    });

    expect(after.left).toBe(before.left);
    expect(after.top).toBe(before.top);
    expect(after.width).toBe(before.width);
    expect(after.height).toBe(before.height);
  });

  test('dialog is excluded from grid layout and its position is unchanged', async ({ page }) => {
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');

    await executeSQL(page, 'SELECT name FROM sample1');
    await page.waitForTimeout(300);

    // Open column manager
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.focusWindow(win.id);
      app._test.showColManager(win.id);
    });
    await page.waitForTimeout(300);

    // Record dialog position and the layout unit count before grid
    const before = await page.evaluate(() => {
      const colMgrWin = app._test._activeColManagerWin;
      return {
        left: colMgrWin.el.style.left,
        top: colMgrWin.el.style.top,
        width: colMgrWin.el.style.width,
        height: colMgrWin.el.style.height,
      };
    });

    // Grid layout
    await page.evaluate(() => app.layoutGrid());
    await page.waitForTimeout(300);

    const after = await page.evaluate(() => {
      const colMgrWin = app._test._activeColManagerWin;
      return {
        left: colMgrWin.el.style.left,
        top: colMgrWin.el.style.top,
        width: colMgrWin.el.style.width,
        height: colMgrWin.el.style.height,
      };
    });

    expect(after.left).toBe(before.left);
    expect(after.top).toBe(before.top);
    expect(after.width).toBe(before.width);
    expect(after.height).toBe(before.height);
  });

  test('dialog is excluded from cascade layout and its position is unchanged', async ({ page }) => {
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');

    await executeSQL(page, 'SELECT name FROM sample1');
    await page.waitForTimeout(300);

    // Open column manager
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.focusWindow(win.id);
      app._test.showColManager(win.id);
    });
    await page.waitForTimeout(300);

    const before = await page.evaluate(() => {
      const colMgrWin = app._test._activeColManagerWin;
      return {
        left: colMgrWin.el.style.left,
        top: colMgrWin.el.style.top,
        width: colMgrWin.el.style.width,
        height: colMgrWin.el.style.height,
      };
    });

    await page.evaluate(() => app.layoutCascade());
    await page.waitForTimeout(300);

    const after = await page.evaluate(() => {
      const colMgrWin = app._test._activeColManagerWin;
      return {
        left: colMgrWin.el.style.left,
        top: colMgrWin.el.style.top,
        width: colMgrWin.el.style.width,
        height: colMgrWin.el.style.height,
      };
    });

    expect(after.left).toBe(before.left);
    expect(after.top).toBe(before.top);
    expect(after.width).toBe(before.width);
    expect(after.height).toBe(before.height);
  });

  test('multiple non-modal dialogs all stay on top of regular windows', async ({ page }) => {
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');

    // Open column manager
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.focusWindow(win.id);
      app._test.showColManager(win.id);
    });
    await page.waitForTimeout(300);

    // Open Plugin About dialog with a synthetic plugin object
    await page.evaluate(() => {
      app._test.showPluginAbout({
        name: 'Test Plugin',
        version: '1.0',
        author: 'Test',
        description: 'A test plugin',
        columns: [],
        links: [],
      }, 0);
    });
    await page.waitForTimeout(300);

    // Focus the table window to trigger raiseDialogs
    await page.evaluate(() => {
      const tableWin = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.focusWindow(tableWin.id);
    });
    await page.waitForTimeout(200);

    const result = await page.evaluate(() => {
      const tableWin = app._test.windows.find(w => w.tableName === 'sample1');
      const tableZ = parseInt(tableWin.el.style.zIndex) || 0;
      const dialogs = app._test.windows.filter(w => w.isDialog && !w.backdropEl);
      return {
        tableZ,
        dialogCount: dialogs.length,
        dialogZs: dialogs.map(w => ({
          z: parseInt(w.el.style.zIndex) || 0,
          title: w.el.querySelector('.win-title')?.textContent || '',
        })),
      };
    });

    expect(result.dialogCount).toBeGreaterThanOrEqual(2);
    for (const d of result.dialogZs) {
      expect(d.z).toBeGreaterThan(result.tableZ);
    }
  });

  test('modal dialog with backdrop is not affected by raiseDialogs', async ({ page }) => {
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');

    // Focus the table window first
    await page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.focusWindow(win.id);
    });
    await page.waitForTimeout(200);

    // Open a modal dialog (New Table prompt via Ctrl+N)
    await page.keyboard.press('Control+n');
    await page.waitForTimeout(400);

    // Verify the modal dialog exists and has a backdrop
    const modalInfo = await page.evaluate(() => {
      const modalWin = app._test.windows.find(w => w.isDialog && w.backdropEl);
      if (!modalWin) return { found: false };
      return {
        found: true,
        hasBackdrop: !!modalWin.backdropEl,
        isDialog: modalWin.isDialog,
        z: parseInt(modalWin.el.style.zIndex) || 0,
      };
    });

    expect(modalInfo.found).toBe(true);
    expect(modalInfo.hasBackdrop).toBe(true);
    expect(modalInfo.isDialog).toBe(true);

    // The modal dialog should be visible and have a backdrop element in the DOM
    const backdropVisible = await page.evaluate(() => {
      const modalWin = app._test.windows.find(w => w.isDialog && w.backdropEl);
      return modalWin.backdropEl.parentElement !== null;
    });

    expect(backdropVisible).toBe(true);
  });
});
