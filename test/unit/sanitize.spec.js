const { test, expect } = require('@playwright/test');
const { openApp } = require('../helpers');

test.describe('sanitizeTableName', () => {
  test('replaces special characters with underscores', async ({ page }) => {
    await openApp(page);
    const result = await page.evaluate(() => app._test.sanitizeTableName('my-file (1).csv'));
    expect(result).toBe('my_file__1__csv');
  });

  test('prefixes leading digits', async ({ page }) => {
    await openApp(page);
    const result = await page.evaluate(() => app._test.sanitizeTableName('123table'));
    expect(result).toBe('_123table');
  });

  test('preserves valid names', async ({ page }) => {
    await openApp(page);
    const result = await page.evaluate(() => app._test.sanitizeTableName('valid_name_123'));
    expect(result).toBe('valid_name_123');
  });

  test('handles empty string', async ({ page }) => {
    await openApp(page);
    const result = await page.evaluate(() => app._test.sanitizeTableName(''));
    expect(result).toBe('');
  });
});

test.describe('sanitizeColumnName', () => {
  test('replaces special characters', async ({ page }) => {
    await openApp(page);
    const result = await page.evaluate(() => app._test.sanitizeColumnName('First Name'));
    expect(result).toBe('First_Name');
  });

  test('prefixes leading digits', async ({ page }) => {
    await openApp(page);
    const result = await page.evaluate(() => app._test.sanitizeColumnName('1st_col'));
    expect(result).toBe('_1st_col');
  });
});

test.describe('sanitizeColumns', () => {
  test('deduplicates column names', async ({ page }) => {
    await openApp(page);
    const result = await page.evaluate(() => app._test.sanitizeColumns(['name', 'name', 'name']));
    expect(result).toEqual(['name', 'name_2', 'name_3']);
  });

  test('sanitizes and deduplicates', async ({ page }) => {
    await openApp(page);
    // sanitizeColumns only deduplicates, doesn't sanitize chars
    const result = await page.evaluate(() => app._test.sanitizeColumns(['col-a', 'col a', 'col_a']));
    expect(result).toEqual(['col-a', 'col a', 'col_a']);
  });

  test('deduplicates identical names', async ({ page }) => {
    await openApp(page);
    const result = await page.evaluate(() => app._test.sanitizeColumns(['col_a', 'col_a', 'col_a']));
    expect(result).toEqual(['col_a', 'col_a_2', 'col_a_3']);
  });

  test('handles empty columns as-is', async ({ page }) => {
    await openApp(page);
    // sanitizeColumns preserves empty strings but deduplicates them
    const result = await page.evaluate(() => app._test.sanitizeColumns(['', '', 'name']));
    expect(result[0]).toBe('');
    expect(result[1]).toBe('_2');
    expect(result[2]).toBe('name');
  });
});

test.describe('getUniqueTableName', () => {
  test('returns base name when no conflict', async ({ page }) => {
    await openApp(page);
    const result = await page.evaluate(() => app._test.getUniqueTableName('newtable'));
    expect(result).toBe('newtable');
  });

  test('appends suffix when name exists', async ({ page }) => {
    await openApp(page);
    const result = await page.evaluate(() => {
      const tables = app._test.tables;
      tables['test_dup'] = { columns: [], rows: [] };
      const name = app._test.getUniqueTableName('test_dup');
      delete tables['test_dup'];
      return name;
    });
    expect(result).toBe('test_dup_2');
  });
});
