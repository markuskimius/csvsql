const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow } = require('../helpers');

test.describe('File menu item order', () => {
  test('New Table is the first item in the File menu', async ({ page }) => {
    await openApp(page);
    await page.locator('#menu-file .menu-label').click();
    const buttons = await page.locator('#menu-file .menu-dropdown button').all();
    const labels = await Promise.all(buttons.map(b => b.textContent()));
    expect(labels[0]).toContain('New Table');
  });

  test('File menu items appear in correct order', async ({ page }) => {
    await openApp(page);
    await page.locator('#menu-file .menu-label').click();
    const buttons = await page.locator('#menu-file .menu-dropdown button').all();
    const labels = await Promise.all(buttons.map(b => b.textContent()));
    expect(labels[0]).toContain('New Table');
    expect(labels[1]).toContain('Open...');
    expect(labels[2]).toContain('Open (no header)');
    expect(labels[3]).toContain('Open URL...');
    expect(labels[4]).toContain('Open URL (no header)');
    expect(labels[5]).toContain('Save Table');
    expect(labels[6]).toContain('Save Table As');
    expect(labels[7]).toContain('Close Window');
  });
});

test.describe('Menu state: table-only items disabled for non-table windows', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
  });

  test('Save and Save As are disabled when no windows are open', async ({ page }) => {
    await page.locator('#menu-file .menu-label').click();
    await expect(page.locator('#btn-save')).toBeDisabled();
    await expect(page.locator('#btn-save-as')).toBeDisabled();
    await expect(page.locator('#btn-close-window')).toBeDisabled();
  });

  test('Save and Save As are enabled when a table window is active', async ({ page }) => {
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
    await page.locator('#menu-file .menu-label').click();
    await expect(page.locator('#btn-save')).toBeEnabled();
    await expect(page.locator('#btn-save-as')).toBeEnabled();
    await expect(page.locator('#btn-close-window')).toBeEnabled();
  });

  test('Save and Save As are disabled when Manual window is active', async ({ page }) => {
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
    await page.evaluate(() => app.showManual());
    await page.waitForSelector('.help-body');
    await page.locator('#menu-file .menu-label').click();
    await expect(page.locator('#btn-save')).toBeDisabled();
    await expect(page.locator('#btn-save-as')).toBeDisabled();
    await expect(page.locator('#btn-close-window')).toBeEnabled();
  });

  test('Save and Save As are disabled when About window is active', async ({ page }) => {
    await page.evaluate(() => app.showAbout());
    await page.waitForSelector('.help-body');
    await page.locator('#menu-file .menu-label').click();
    await expect(page.locator('#btn-save')).toBeDisabled();
    await expect(page.locator('#btn-save-as')).toBeDisabled();
    await expect(page.locator('#btn-close-window')).toBeEnabled();
  });

  test('Save and Save As are disabled when Expression Reference is active', async ({ page }) => {
    await page.evaluate(() => app.showExpressionReference());
    await page.waitForSelector('.help-body');
    await page.locator('#menu-file .menu-label').click();
    await expect(page.locator('#btn-save')).toBeDisabled();
    await expect(page.locator('#btn-save-as')).toBeDisabled();
    await expect(page.locator('#btn-close-window')).toBeEnabled();
  });

  test('Save and Save As re-enable when switching back to a table window', async ({ page }) => {
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');

    // Open Manual — it becomes the active window, Save should be disabled
    await page.evaluate(() => app.showManual());
    await page.waitForSelector('.help-body');
    await page.locator('#menu-file .menu-label').click();
    await expect(page.locator('#btn-save')).toBeDisabled();
    await page.locator('#menu-file .menu-label').click(); // close menu

    // Close Manual via its close button — table window regains focus
    const closeBtn = page.locator('.subwindow.active .btn-close');
    await closeBtn.click();
    await page.waitForTimeout(100);

    await page.locator('#menu-file .menu-label').click();
    await expect(page.locator('#btn-save')).toBeEnabled();
    await expect(page.locator('#btn-save-as')).toBeEnabled();
  });

  test('Edit menu items are disabled when a non-table window is active', async ({ page }) => {
    await page.evaluate(() => app.showManual());
    await page.waitForSelector('.help-body');
    await page.locator('#menu-edit .menu-label').click();
    await expect(page.locator('#btn-undo')).toBeDisabled();
    await expect(page.locator('#btn-redo')).toBeDisabled();
    await expect(page.locator('#btn-cut')).toBeDisabled();
    await expect(page.locator('#btn-copy')).toBeDisabled();
    await expect(page.locator('#btn-paste')).toBeDisabled();
    await expect(page.locator('#btn-select-all')).toBeDisabled();
    await expect(page.locator('#btn-select-none')).toBeDisabled();
    await expect(page.locator('#btn-col-manager')).toBeDisabled();
  });

  test('Close Window is enabled for any window type', async ({ page }) => {
    await page.evaluate(() => app.showManual());
    await page.waitForSelector('.help-body');
    await page.locator('#menu-file .menu-label').click();
    await expect(page.locator('#btn-close-window')).toBeEnabled();
  });

  test('Windows layout buttons are enabled regardless of active window type', async ({ page }) => {
    await page.evaluate(() => app.showManual());
    await page.waitForSelector('.help-body');
    await page.locator('#menu-windows .menu-label').click();
    await expect(page.locator('#btn-tile-h')).toBeEnabled();
    await expect(page.locator('#btn-tile-v')).toBeEnabled();
    await expect(page.locator('#btn-grid')).toBeEnabled();
    await expect(page.locator('#btn-cascade')).toBeEnabled();
  });
});
