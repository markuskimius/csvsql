const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, executeSQL } = require('../helpers');

const Q1 = 'SELECT name FROM sample1';
const Q2 = 'SELECT email FROM sample1 LIMIT 3';

const input = (page) => page.locator('#sql-input');

test.describe('SQL query history', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('ArrowUp recalls previous queries, ArrowDown returns to the draft', async ({ page }) => {
    await executeSQL(page, Q1);
    await executeSQL(page, Q2);
    await input(page).fill('');
    await input(page).press('ArrowUp');
    await expect(input(page)).toHaveValue(Q2);
    await input(page).press('ArrowUp');
    await expect(input(page)).toHaveValue(Q1);
    // Down (with the cursor at the end) walks forward; past the newest entry
    // it restores the (empty) draft
    await input(page).press('End');
    await input(page).press('ArrowDown');
    await expect(input(page)).toHaveValue(Q2);
    await input(page).press('ArrowDown');
    await expect(input(page)).toHaveValue('');
  });

  test('ArrowUp preserves an unsubmitted draft', async ({ page }) => {
    await executeSQL(page, Q1);
    await input(page).fill('SELECT draft');
    // Cursor is at the end after fill — move it to the start for history recall
    await input(page).press('Home');
    await input(page).press('ArrowUp');
    await expect(input(page)).toHaveValue(Q1);
    await input(page).press('End');
    await input(page).press('ArrowDown');
    await expect(input(page)).toHaveValue('SELECT draft');
  });

  test('re-running a query moves it to the front without duplicating', async ({ page }) => {
    await executeSQL(page, Q1);
    await executeSQL(page, Q2);
    await executeSQL(page, Q1);
    await input(page).fill('');
    await input(page).press('ArrowUp');
    await expect(input(page)).toHaveValue(Q1);
    await input(page).press('ArrowUp');
    await expect(input(page)).toHaveValue(Q2);
    // Only two entries — another ArrowUp stays on the oldest
    await input(page).press('ArrowUp');
    await expect(input(page)).toHaveValue(Q2);
  });

  test('history persists across page reloads', async ({ page }) => {
    await executeSQL(page, Q1);
    await page.reload();
    await page.waitForFunction(() => window._appReady === true, undefined, { timeout: 15000 });
    await input(page).click();
    await input(page).press('ArrowUp');
    await expect(input(page)).toHaveValue(Q1);
  });

  test('recall updates the syntax-highlight overlay', async ({ page }) => {
    await executeSQL(page, Q1);
    await input(page).fill('');
    await input(page).press('ArrowUp');
    await expect(page.locator('#sql-highlight')).toContainText('SELECT name FROM sample1');
  });

  test('ArrowUp mid-text does not recall history', async ({ page }) => {
    await executeSQL(page, Q1);
    await input(page).fill('SELECT 1');
    // Cursor at the end; ArrowUp requires the cursor at the very start
    await input(page).press('ArrowUp');
    await expect(input(page)).toHaveValue('SELECT 1');
  });

  test('arrows navigate the autocomplete dropdown, not history, while it is open', async ({ page }) => {
    await executeSQL(page, Q1);
    // Enter history (idx now points at Q1), then edit at the end so an
    // ArrowDown WOULD move forward in history if the dropdown guard failed
    await input(page).fill('');
    await input(page).press('ArrowUp');
    await expect(input(page)).toHaveValue(Q1);
    await input(page).press('End');
    await input(page).pressSequentially(' LIM'); // opens the dropdown (LIMIT)
    await expect(page.locator('#sql-autocomplete')).toBeVisible();
    const before = await input(page).inputValue();
    await input(page).press('ArrowDown');
    expect(await input(page).inputValue()).toBe(before);
    await expect(page.locator('#sql-autocomplete')).toBeVisible();
  });
});
