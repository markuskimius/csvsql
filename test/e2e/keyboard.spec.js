const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow } = require('../helpers');

test.describe('Keyboard Shortcuts', () => {
  test('Ctrl+N creates a new table', async ({ page }) => {
    await openApp(page);

    await page.keyboard.press('Control+n');

    // Custom modal prompt appears — fill in table name
    const tableNameInput = page.locator('.modal-input');
    await expect(tableNameInput).toBeVisible();
    await tableNameInput.fill('keyboard_test');
    await page.locator('.modal .ok').click();

    // Second prompt for column names
    const colInput = page.locator('.modal-input');
    await expect(colInput).toBeVisible();
    await colInput.fill('id, name');
    await page.locator('.modal .ok').click();

    await waitForWindow(page, 'keyboard_test');
    const count = await page.evaluate(() => app._test.windows.length);
    expect(count).toBe(1);
  });

  test('Ctrl+Enter executes SQL query', async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');

    await page.locator('#sql-input').fill('SELECT * FROM [sample1]');
    await page.locator('#sql-input').press('Control+Enter');

    // Wait for a result to appear
    for (let i = 0; i < 25; i++) {
      const count = await page.evaluate(() => app._test.windows.length);
      if (count >= 2) break;
      await page.waitForTimeout(200);
    }
    const count = await page.evaluate(() => app._test.windows.length);
    expect(count).toBe(2);
  });

  test('Ctrl+W closes active window', async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');

    // Accept any confirmation dialog
    page.on('dialog', dialog => dialog.accept());

    await page.keyboard.press('Control+w');
    await page.waitForTimeout(500);

    const count = await page.evaluate(() => app._test.windows.length);
    expect(count).toBe(0);
  });
});
