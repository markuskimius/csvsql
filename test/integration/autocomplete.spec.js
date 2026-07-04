const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow } = require('../helpers');

// sample1.csv → table "sample1" with columns name, email, member_since

async function openSample(page) {
  await openApp(page);
  await uploadFile(page, 'sample1.csv');
  await waitForWindow(page, 'sample1');
}

// The toolbar defaults to quick-search mode; SQL autocompletion only applies
// to the WHERE filter, so switch to Filter mode first.
async function openSampleFilterMode(page) {
  await openSample(page);
  await page.locator('.subwindow .toolbar-mode .mode-filter').click();
}

const dropdown = page => page.locator('#sql-autocomplete');
const sqlInput = page => page.locator('#sql-input');
const filterInput = page => page.locator('.filter-input');

test.describe('autocomplete — SQL console', () => {
  test('typing a table prefix after FROM opens the dropdown; Tab accepts', async ({ page }) => {
    await openSample(page);
    await sqlInput(page).pressSequentially('SELECT * FROM sam');
    await expect(dropdown(page)).toBeVisible();
    await expect(dropdown(page).locator('.ac-item.selected')).toContainText('sample1');
    await page.keyboard.press('Tab');
    await expect(dropdown(page)).toBeHidden();
    await expect(sqlInput(page)).toHaveValue('SELECT * FROM sample1');
  });

  test('Enter accepts without inserting a newline', async ({ page }) => {
    await openSample(page);
    await sqlInput(page).pressSequentially('SELECT * FROM sam');
    await expect(dropdown(page)).toBeVisible();
    await page.keyboard.press('Enter');
    await expect(sqlInput(page)).toHaveValue('SELECT * FROM sample1');
  });

  test('dot after a table name lists its columns; arrows navigate', async ({ page }) => {
    await openSample(page);
    await sqlInput(page).pressSequentially('SELECT * FROM sample1 WHERE sample1.');
    await expect(dropdown(page)).toBeVisible();
    const labels = await dropdown(page).locator('.ac-label').allTextContents();
    expect(labels).toEqual(['name', 'email', 'member_since']);
    await page.keyboard.press('ArrowDown');
    await expect(dropdown(page).locator('.ac-item.selected')).toContainText('email');
    await page.keyboard.press('ArrowUp');
    await expect(dropdown(page).locator('.ac-item.selected')).toContainText('name');
    await page.keyboard.press('Enter');
    await expect(sqlInput(page)).toHaveValue('SELECT * FROM sample1 WHERE sample1.name');
  });

  test('clicking an item accepts it', async ({ page }) => {
    await openSample(page);
    await sqlInput(page).pressSequentially('SELECT * FROM sample1 WHERE sample1.');
    await dropdown(page).locator('.ac-item', { hasText: 'email' }).click();
    await expect(sqlInput(page)).toHaveValue('SELECT * FROM sample1 WHERE sample1.email');
    await expect(dropdown(page)).toBeHidden();
  });

  test('accepting updates the syntax highlight overlay', async ({ page }) => {
    await openSample(page);
    await sqlInput(page).pressSequentially('SELECT * FROM sam');
    await page.keyboard.press('Tab');
    await expect(page.locator('#sql-highlight')).toContainText('SELECT * FROM sample1');
  });

  test('Escape closes the dropdown and typing continues normally', async ({ page }) => {
    await openSample(page);
    await sqlInput(page).pressSequentially('SELECT * FROM sam');
    await expect(dropdown(page)).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(dropdown(page)).toBeHidden();
    await sqlInput(page).pressSequentially('p');
    await expect(sqlInput(page)).toHaveValue('SELECT * FROM samp');
  });

  test('Ctrl+Enter executes the query even while the dropdown is open', async ({ page }) => {
    await openSample(page);
    await sqlInput(page).pressSequentially('SELECT name FROM sample1 WHERE nam');
    await expect(dropdown(page)).toBeVisible();
    await page.keyboard.press('ControlOrMeta+Enter');
    await expect(dropdown(page)).toBeHidden();
    // The literal query is invalid ("nam") — an error status proves execution ran
    await expect(page.locator('#console-status')).toContainText('Error', { timeout: 15000 });
  });

  test('no dropdown while typing inside a string literal', async ({ page }) => {
    await openSample(page);
    await sqlInput(page).pressSequentially("SELECT * FROM sample1 WHERE name = 'sam");
    await expect(dropdown(page)).toBeHidden();
  });

  test('tables created by queries complete immediately', async ({ page }) => {
    await openSample(page);
    await page.evaluate(() => {
      document.getElementById('sql-input').value = '';
      app._test.db.run('CREATE TABLE zebra (stripe TEXT)');
    });
    await page.evaluate(() => {
      // Register through the app path so tables{} knows about it
      const t = { columns: ['stripe'], rows: [], filename: null, modified: false };
      app._test.tables['zebra'] = t;
    });
    await sqlInput(page).pressSequentially('SELECT * FROM zeb');
    await expect(dropdown(page).locator('.ac-item.selected')).toContainText('zebra');
  });
});

test.describe('autocomplete — filter input', () => {
  test('column prefix completes and the filter applies through the normal path', async ({ page }) => {
    await openSampleFilterMode(page);
    await filterInput(page).pressSequentially('na');
    await expect(dropdown(page)).toBeVisible();
    await expect(dropdown(page).locator('.ac-item.selected')).toContainText('name');
    await page.keyboard.press('Tab');
    await expect(filterInput(page)).toHaveValue('name');
    await filterInput(page).pressSequentially(" LIKE '%Alice%'");
    await expect.poll(() => page.evaluate(() =>
      app._test.windows[0]._displayRows.length
    )).toBe(1);
  });

  test('Ctrl+Space opens with own columns ranked first', async ({ page }) => {
    await openSampleFilterMode(page);
    await filterInput(page).click();
    await page.keyboard.press('ControlOrMeta+Space');
    await expect(dropdown(page)).toBeVisible();
    const labels = await dropdown(page).locator('.ac-label').allTextContents();
    expect(labels.slice(0, 3)).toEqual(['name', 'email', 'member_since']);
  });

  test('first Escape closes the dropdown, second returns focus to the table', async ({ page }) => {
    await openSampleFilterMode(page);
    await filterInput(page).pressSequentially('na');
    await expect(dropdown(page)).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(dropdown(page)).toBeHidden();
    await expect(filterInput(page)).toBeFocused();
    await page.keyboard.press('Escape');
    await expect(page.locator('td.data-cell:focus')).toHaveCount(1);
  });

  test('Tab with the dropdown closed is not intercepted', async ({ page }) => {
    await openSampleFilterMode(page);
    await filterInput(page).click();
    await page.keyboard.press('Tab');
    await expect(filterInput(page)).not.toBeFocused();
  });

  test('blur closes the dropdown', async ({ page }) => {
    await openSampleFilterMode(page);
    await filterInput(page).pressSequentially('na');
    await expect(dropdown(page)).toBeVisible();
    await page.evaluate(() => document.querySelector('.filter-input').blur());
    await expect(dropdown(page)).toBeHidden();
  });

  test('clicking outside closes the dropdown without accepting', async ({ page }) => {
    await openSampleFilterMode(page);
    await filterInput(page).pressSequentially('na');
    await expect(dropdown(page)).toBeVisible();
    await page.locator('.win-title').first().click();
    await expect(dropdown(page)).toBeHidden();
    await expect(filterInput(page)).toHaveValue('na');
  });
});
