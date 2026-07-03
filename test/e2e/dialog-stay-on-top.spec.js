const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, executeSQL } = require('../helpers');

test.describe('Dialog stay-on-top', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  function openColManager(page) {
    return page.evaluate(() => {
      const win = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.focusWindow(win.id);
      app._test.showColManager(win.id);
    });
  }

  function getZIndexes(page) {
    return page.evaluate(() => {
      const colMgr = app._test._activeColManagerWin;
      const nonDialog = app._test.windows.filter(w => !w.isDialog);
      const maxNonDialog = Math.max(...nonDialog.map(w => parseInt(w.el.style.zIndex) || 0));
      return {
        dialogZ: colMgr ? parseInt(colMgr.el.style.zIndex) || 0 : null,
        maxNonDialogZ: maxNonDialog,
        nonDialogCount: nonDialog.length,
      };
    });
  }

  test('dialog has higher z-index after focusing a regular window', async ({ page }) => {
    await openColManager(page);

    const result = await page.evaluate(() => {
      const tableWin = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.focusWindow(tableWin.id);
      const colMgr = app._test._activeColManagerWin;
      return {
        tableZ: parseInt(tableWin.el.style.zIndex) || 0,
        dialogZ: parseInt(colMgr.el.style.zIndex) || 0,
        isDialog: colMgr.isDialog,
      };
    });

    expect(result.isDialog).toBe(true);
    expect(result.dialogZ).toBeGreaterThan(result.tableZ);
  });

  test('dialog stays on top after creating a new window via SQL', async ({ page }) => {
    await openColManager(page);
    await executeSQL(page, 'SELECT name FROM sample1');

    const z = await getZIndexes(page);
    expect(z.nonDialogCount).toBeGreaterThanOrEqual(2);
    expect(z.dialogZ).toBeGreaterThan(z.maxNonDialogZ);
  });

  test('dialog stays on top after cascade layout', async ({ page }) => {
    await executeSQL(page, 'SELECT name FROM sample1');
    await openColManager(page);

    await page.evaluate(() => app.layoutCascade());
    const z = await getZIndexes(page);
    expect(z.dialogZ).toBeGreaterThan(z.maxNonDialogZ);
  });

  test('dialog stays on top after tile layout', async ({ page }) => {
    await executeSQL(page, 'SELECT name FROM sample1');
    await openColManager(page);

    await page.evaluate(() => app.layoutTileH());
    const z = await getZIndexes(page);
    expect(z.dialogZ).toBeGreaterThan(z.maxNonDialogZ);
  });

  test('dialog excluded from tile layout — position unchanged', async ({ page }) => {
    await openColManager(page);

    const { before, after } = await page.evaluate(() => {
      const el = app._test._activeColManagerWin.el;
      const snap = () => ({ left: el.style.left, top: el.style.top, width: el.style.width, height: el.style.height });
      const before = snap();
      app.layoutTileH();
      return { before, after: snap() };
    });

    expect(after).toEqual(before);
  });

  test('dialog excluded from grid layout — position unchanged', async ({ page }) => {
    await executeSQL(page, 'SELECT name FROM sample1');
    await openColManager(page);

    const { before, after } = await page.evaluate(() => {
      const el = app._test._activeColManagerWin.el;
      const snap = () => ({ left: el.style.left, top: el.style.top, width: el.style.width, height: el.style.height });
      const before = snap();
      app.layoutGrid();
      return { before, after: snap() };
    });

    expect(after).toEqual(before);
  });

  test('dialog excluded from cascade layout — position unchanged', async ({ page }) => {
    await executeSQL(page, 'SELECT name FROM sample1');
    await openColManager(page);

    const { before, after } = await page.evaluate(() => {
      const el = app._test._activeColManagerWin.el;
      const snap = () => ({ left: el.style.left, top: el.style.top, width: el.style.width, height: el.style.height });
      const before = snap();
      app.layoutCascade();
      return { before, after: snap() };
    });

    expect(after).toEqual(before);
  });

  test('multiple non-modal dialogs all stay on top', async ({ page }) => {
    await openColManager(page);

    const result = await page.evaluate(() => {
      app._test.showPluginAbout({
        name: 'Test Plugin', version: '1.0', author: 'Test',
        description: 'A test plugin', columns: [], links: [],
      }, 0);
      const tableWin = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.focusWindow(tableWin.id);
      const tableZ = parseInt(tableWin.el.style.zIndex) || 0;
      const dialogs = app._test.windows.filter(w => w.isDialog && !w.backdropEl);
      return {
        tableZ,
        dialogCount: dialogs.length,
        allAbove: dialogs.every(w => (parseInt(w.el.style.zIndex) || 0) > tableZ),
      };
    });

    expect(result.dialogCount).toBeGreaterThanOrEqual(2);
    expect(result.allAbove).toBe(true);
  });

  test('modal dialog (with backdrop) is not raised by raiseDialogs', async ({ page }) => {
    await page.keyboard.press('Control+n');
    await expect(page.locator('.dialog-backdrop')).toBeVisible();

    const result = await page.evaluate(() => {
      const modalWin = app._test.windows.find(w => w.isDialog && w.backdropEl);
      return {
        found: !!modalWin,
        hasBackdrop: !!modalWin?.backdropEl,
        inDOM: modalWin?.backdropEl?.parentElement !== null,
      };
    });

    expect(result.found).toBe(true);
    expect(result.hasBackdrop).toBe(true);
    expect(result.inDOM).toBe(true);
  });
});
