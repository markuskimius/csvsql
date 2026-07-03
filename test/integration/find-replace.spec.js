const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, getTableData } = require('../helpers');

// Debounce on the find input is 150ms — wait a bit longer after typing.
const DEBOUNCE = 350;

test.describe('Find & Replace', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('Ctrl+F opens the dialog targeting the active table', async ({ page }) => {
    await page.keyboard.press('Control+f');
    await expect(page.locator('.find-input')).toBeVisible();
    await expect(page.locator('.subwindow .win-title', { hasText: 'Find & Replace — sample1' })).toBeVisible();
    // Find input has focus
    await expect(page.locator('.find-input')).toBeFocused();
  });

  test('Edit menu item opens the dialog', async ({ page }) => {
    await page.locator('#menu-edit .menu-label').click();
    await page.locator('#btn-find').click();
    await expect(page.locator('.find-input')).toBeVisible();
  });

  test('typing a query highlights matches and shows the count', async ({ page }) => {
    await page.keyboard.press('Control+f');
    await page.locator('.find-input').fill('example.com');
    await page.waitForTimeout(DEBOUNCE);
    await expect(page.locator('.find-count')).toHaveText('1 of 10');
    await expect(page.locator('.subwindow table tbody td.cell-find-match')).toHaveCount(10);
    await expect(page.locator('.subwindow table tbody td.cell-find-current')).toHaveCount(1);
  });

  test('Next and Prev navigate between matches', async ({ page }) => {
    await page.keyboard.press('Control+f');
    await page.locator('.find-input').fill('example.com');
    await page.waitForTimeout(DEBOUNCE);
    await page.locator('.find-next').click();
    await expect(page.locator('.find-count')).toHaveText('2 of 10');
    await page.locator('.find-prev').click();
    await expect(page.locator('.find-count')).toHaveText('1 of 10');
    // Prev from the first match wraps to the last
    await page.locator('.find-prev').click();
    await expect(page.locator('.find-count')).toHaveText('10 of 10');
  });

  test('Enter and Shift+Enter in the find input navigate', async ({ page }) => {
    await page.keyboard.press('Control+f');
    await page.locator('.find-input').fill('example.com');
    await page.waitForTimeout(DEBOUNCE);
    await page.locator('.find-input').press('Enter');
    await expect(page.locator('.find-count')).toHaveText('2 of 10');
    await page.locator('.find-input').press('Shift+Enter');
    await expect(page.locator('.find-count')).toHaveText('1 of 10');
  });

  test('match case option filters matches', async ({ page }) => {
    await page.keyboard.press('Control+f');
    await page.locator('.find-input').fill('alice');
    await page.waitForTimeout(DEBOUNCE);
    // Case-insensitive: matches "Alice Johnson" and "alice@example.com"
    await expect(page.locator('.find-count')).toHaveText('1 of 2');
    await page.locator('.find-case').check();
    await expect(page.locator('.find-count')).toHaveText('1 of 1');
  });

  test('entire cell option requires a full-cell match', async ({ page }) => {
    await page.keyboard.press('Control+f');
    await page.locator('.find-input').fill('Bob');
    await page.waitForTimeout(DEBOUNCE);
    // Case-insensitive: matches "Bob Smith" and "bob@example.com"
    await expect(page.locator('.find-count')).toHaveText('1 of 2');
    await page.locator('.find-whole').check();
    await expect(page.locator('.find-count')).toHaveText('No matches');
    await page.locator('.find-input').fill('Bob Smith');
    await page.waitForTimeout(DEBOUNCE);
    await expect(page.locator('.find-count')).toHaveText('1 of 1');
  });

  test('Replace replaces the current match', async ({ page }) => {
    await page.keyboard.press('Control+f');
    await page.locator('.find-input').fill('Alice');
    await page.waitForTimeout(DEBOUNCE);
    await page.locator('.find-replace-input').fill('Alicia');
    // First match is already selected by the auto-goto, so one click replaces it
    await page.locator('.find-replace-one').click();
    await page.waitForTimeout(200);
    const data = await getTableData(page, 'sample1');
    expect(data.rows[0].name).toBe('Alicia Johnson');
    expect(data.modified).toBe(true);
  });

  test('Replace All replaces every match as a single undo entry', async ({ page }) => {
    await page.keyboard.press('Control+f');
    await page.locator('.find-input').fill('example.com');
    await page.waitForTimeout(DEBOUNCE);
    await page.locator('.find-replace-input').fill('test.org');
    await page.locator('.find-replace-all').click();
    await page.waitForTimeout(200);

    await expect(page.locator('.toast', { hasText: 'Replaced in 10 cells' })).toBeVisible();
    let data = await getTableData(page, 'sample1');
    expect(data.rows.every(r => r.email.endsWith('@test.org'))).toBe(true);
    await expect(page.locator('.find-count')).toHaveText('No matches');

    // One undo restores all 10 cells
    await page.evaluate(() => document.querySelectorAll('.toast').forEach(t => t.remove()));
    await page.keyboard.press('Control+z');
    await page.waitForTimeout(200);
    await expect(page.locator('.toast', { hasText: 'Undid Replace (10 cells)' })).toBeVisible();
    data = await getTableData(page, 'sample1');
    expect(data.rows.every(r => r.email.endsWith('@example.com'))).toBe(true);
  });

  test('Escape closes the dialog and clears highlights', async ({ page }) => {
    await page.keyboard.press('Control+f');
    await page.locator('.find-input').fill('example.com');
    await page.waitForTimeout(DEBOUNCE);
    await expect(page.locator('.subwindow table tbody td.cell-find-match')).toHaveCount(10);
    await page.locator('.find-input').press('Escape');
    await expect(page.locator('.find-input')).toHaveCount(0);
    await expect(page.locator('.subwindow table tbody td.cell-find-match')).toHaveCount(0);
  });

  test('closing the target window closes the dialog', async ({ page }) => {
    await page.keyboard.press('Control+f');
    await expect(page.locator('.find-input')).toBeVisible();
    await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'sample1');
      app._test.tables.sample1.modified = false; // skip the unsaved-changes prompt
      app._test.closeWindow(w.id);
    });
    await expect(page.locator('.find-input')).toHaveCount(0);
  });

  test('navigation scrolls an off-screen match into view', async ({ page }) => {
    // Shrink the window so only a few rows are visible
    await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'sample1');
      w.el.style.height = '160px';
      w.el.querySelector('.table-container').dispatchEvent(new Event('scroll'));
    });
    await page.keyboard.press('Control+f');
    await page.locator('.find-input').fill('Jack Thomas'); // last row
    await page.waitForTimeout(DEBOUNCE);
    await expect(page.locator('.find-count')).toHaveText('1 of 1');
    const current = page.locator('.subwindow table tbody td.cell-find-current');
    await expect(current).toHaveCount(1);
    await expect(current).toBeInViewport();
  });
});
