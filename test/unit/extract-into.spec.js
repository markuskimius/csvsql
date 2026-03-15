const { test, expect } = require('@playwright/test');
const { openApp } = require('../helpers');

test.describe('extractIntoClause', () => {
  test('extracts INTO from SELECT INTO statement', async ({ page }) => {
    await openApp(page);
    const result = await page.evaluate(() =>
      app._test.extractIntoClause('SELECT * INTO newtable FROM old')
    );
    expect(result).not.toBeNull();
    expect(result.targetName).toBe('newtable');
    expect(result.selectSQL).toContain('SELECT *');
    expect(result.selectSQL).toContain('FROM old');
    expect(result.selectSQL).not.toContain('INTO');
  });

  test('handles bracket-quoted table name', async ({ page }) => {
    await openApp(page);
    const result = await page.evaluate(() =>
      app._test.extractIntoClause('SELECT col1 INTO [mytable] FROM src')
    );
    expect(result).not.toBeNull();
    expect(result.targetName).toBe('mytable');
  });

  test('returns null for non-SELECT statements', async ({ page }) => {
    await openApp(page);
    const result = await page.evaluate(() =>
      app._test.extractIntoClause('INSERT INTO tablename VALUES (1)')
    );
    expect(result).toBeNull();
  });

  test('returns null when no INTO clause', async ({ page }) => {
    await openApp(page);
    const result = await page.evaluate(() =>
      app._test.extractIntoClause('SELECT * FROM tablename')
    );
    expect(result).toBeNull();
  });

  test('handles case insensitivity', async ({ page }) => {
    await openApp(page);
    const result = await page.evaluate(() =>
      app._test.extractIntoClause('select * into MyTable from source')
    );
    expect(result).not.toBeNull();
    expect(result.targetName).toBe('MyTable');
  });
});
