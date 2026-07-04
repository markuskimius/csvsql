const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, executeSQL, getWindowCount } = require('../helpers');

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

test.describe('SQL history panel', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('History button lists queries newest-first; clicking an entry loads it for editing', async ({ page }) => {
    await executeSQL(page, Q1);
    await executeSQL(page, Q2);
    await input(page).fill('');
    await page.locator('#btn-sql-history').click();
    const items = page.locator('.sql-history-item');
    await expect(items).toHaveCount(2);
    await expect(items.nth(0)).toContainText(Q2);
    await expect(items.nth(1)).toContainText(Q1);
    await items.nth(1).click();
    // Panel closes; the entry is in the input, focused, with the overlay synced
    await expect(page.locator('#sql-history-panel')).toHaveCount(0);
    await expect(input(page)).toHaveValue(Q1);
    await expect(input(page)).toBeFocused();
    await expect(page.locator('#sql-highlight')).toContainText(Q1);
  });

  test('the run button executes the entry immediately', async ({ page }) => {
    await executeSQL(page, Q1);
    const before = await getWindowCount(page);
    await input(page).fill('');
    await page.locator('#btn-sql-history').click();
    await page.locator('.sql-history-item .sql-history-run').first().click();
    await expect(page.locator('#sql-history-panel')).toHaveCount(0);
    // A new result window opens from the executed query
    await page.waitForFunction((n) => app._test.windows.length > n, before);
  });

  test('Clear History erases the history', async ({ page }) => {
    await executeSQL(page, Q1);
    await input(page).fill('');
    await page.locator('#btn-sql-history').click();
    await page.locator('.sql-history-clear').click();
    await expect(page.locator('.sql-history-empty')).toBeVisible();
    await expect(page.locator('.sql-history-clear')).toBeDisabled();
    const stored = await page.evaluate(() => localStorage.getItem('csvsql_sql_history'));
    expect(JSON.parse(stored)).toEqual([]);
    // ArrowUp no longer recalls anything
    await page.keyboard.press('Escape');
    await input(page).click();
    await input(page).press('ArrowUp');
    await expect(input(page)).toHaveValue('');
  });

  test('panel toggles via the button and closes on Escape or outside click', async ({ page }) => {
    await executeSQL(page, Q1);
    const btn = page.locator('#btn-sql-history');
    const panel = page.locator('#sql-history-panel');
    await btn.click();
    await expect(panel).toBeVisible();
    await btn.click();
    await expect(panel).toHaveCount(0);
    await btn.click();
    await expect(panel).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(panel).toHaveCount(0);
    await btn.click();
    await expect(panel).toBeVisible();
    await page.locator('#menubar').click();
    await expect(panel).toHaveCount(0);
  });

  test('panel fits a short narrow viewport and survives resize', async ({ page }) => {
    await executeSQL(page, Q1);
    await page.setViewportSize({ width: 640, height: 380 });
    await page.locator('#btn-sql-history').click();
    const panel = page.locator('#sql-history-panel');
    await expect(panel).toBeVisible();
    let box = await panel.boundingBox();
    expect(box.y).toBeGreaterThanOrEqual(0);
    expect(box.x).toBeGreaterThanOrEqual(0);
    expect(box.x + box.width).toBeLessThanOrEqual(640);
    // A viewport resize (mobile soft keyboards fire these) repositions the
    // panel instead of closing it
    await page.setViewportSize({ width: 500, height: 300 });
    await expect(panel).toBeVisible();
    box = await panel.boundingBox();
    expect(box.y).toBeGreaterThanOrEqual(0);
    expect(box.x).toBeGreaterThanOrEqual(0);
    expect(box.x + box.width).toBeLessThanOrEqual(500);
    expect(box.y + box.height).toBeLessThanOrEqual(300);
  });

  test('empty history shows a placeholder', async ({ page }) => {
    await page.locator('#btn-sql-history').click();
    await expect(page.locator('.sql-history-empty')).toHaveText('No queries yet');
    await expect(page.locator('.sql-history-clear')).toBeDisabled();
  });
});

test.describe('SQL history on touch devices', () => {
  test.use({ hasTouch: true, isMobile: true });

  test('run button executes without focusing the input (no soft keyboard)', async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
    await executeSQL(page, Q1);
    await input(page).fill('');
    await input(page).blur();
    const before = await getWindowCount(page);
    await page.locator('#btn-sql-history').click();
    await page.locator('.sql-history-item .sql-history-run').first().click();
    await expect(page.locator('#sql-history-panel')).toHaveCount(0);
    await page.waitForFunction((n) => app._test.windows.length > n, before);
    await expect(input(page)).toHaveValue(Q1);
    await expect(input(page)).not.toBeFocused();
  });
});
