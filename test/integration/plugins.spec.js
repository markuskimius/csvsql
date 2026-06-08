const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, getTableData } = require('../helpers');
const path = require('path');

test.describe('Plugin system', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, 'sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  async function loadPluginConfig(page, config) {
    await page.evaluate((cfg) => {
      const errors = app._test.validatePlugin(cfg);
      if (errors.length) throw new Error(errors.join(', '));
      const compiled = app._test.compilePlugin(cfg);
      cfg._compiled = compiled;
      app._test.plugins.push(cfg);
      app._test.rebuildTransformCache();
    }, config);
  }

  async function getCellText(page, rowIdx, colIdx) {
    return page.evaluate(({ rowIdx, colIdx }) => {
      const cells = document.querySelectorAll('.subwindow table tbody tr:not(.virtual-pad)');
      if (!cells[rowIdx]) return null;
      const tds = cells[rowIdx].querySelectorAll('td.data-cell');
      return tds[colIdx] ? tds[colIdx].textContent : null;
    }, { rowIdx, colIdx });
  }

  async function rebuildAndGetCell(page, rowIdx, colIdx) {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      if (win && win._scrollContainer) {
        win._scrollContainer.dispatchEvent(new Event('scroll'));
      }
    });
    await page.waitForTimeout(100);
    return getCellText(page, rowIdx, colIdx);
  }

  test('display transform formats cell values', async ({ page }) => {
    const before = await getCellText(page, 0, 0);
    expect(before).toBe('Alice Johnson');

    await loadPluginConfig(page, {
      name: 'Upper Names',
      table: '.*',
      columns: [{ match: '^name$', display: "upper(value)" }]
    });

    await page.evaluate(() => {
      const win = app._test.windows[0];
      if (win && win.tableName) {
        const body = win.el.querySelector('.win-body');
        const container = win.el.querySelector('.table-container');
        if (container) container.dispatchEvent(new Event('scroll'));
      }
    });
    await page.waitForTimeout(200);

    const displayed = await page.evaluate(() => {
      const win = app._test.windows[0];
      return app._test.getDisplayValue(win.tableName, 'name', app._test.tables[win.tableName].rows[0]);
    });
    expect(displayed).toBe('ALICE JOHNSON');
  });

  test('getDisplayValue returns raw value when no plugin matches', async ({ page }) => {
    const result = await page.evaluate(() => {
      const win = app._test.windows[0];
      const row = app._test.tables[win.tableName].rows[0];
      return app._test.getDisplayValue(win.tableName, 'name', row);
    });
    expect(result).toBe('Alice Johnson');
  });

  test('hasDisplayTransform reflects plugin state', async ({ page }) => {
    const before = await page.evaluate(() => {
      const win = app._test.windows[0];
      return app._test.hasDisplayTransform(win.tableName, 'name');
    });
    expect(before).toBe(false);

    await loadPluginConfig(page, {
      name: 'Test',
      table: '.*',
      columns: [{ match: '^name$', display: "upper(value)" }]
    });

    const after = await page.evaluate(() => {
      const win = app._test.windows[0];
      return app._test.hasDisplayTransform(win.tableName, 'name');
    });
    expect(after).toBe(true);
  });

  test('unloading a plugin reverts display', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Upper Names',
      table: '.*',
      columns: [{ match: '^name$', display: "upper(value)" }]
    });

    let result = await page.evaluate(() => {
      const win = app._test.windows[0];
      return app._test.getDisplayValue(win.tableName, 'name', app._test.tables[win.tableName].rows[0]);
    });
    expect(result).toBe('ALICE JOHNSON');

    await page.evaluate(() => {
      app._test.unloadPlugin(0);
    });

    result = await page.evaluate(() => {
      const win = app._test.windows[0];
      return app._test.getDisplayValue(win.tableName, 'name', app._test.tables[win.tableName].rows[0]);
    });
    expect(result).toBe('Alice Johnson');
  });

  test('multiple plugins stack on different columns', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Upper Names',
      table: '.*',
      columns: [{ match: '^name$', display: "upper(value)" }]
    });
    await loadPluginConfig(page, {
      name: 'Lower Emails',
      table: '.*',
      columns: [{ match: '^email$', display: "lower(value)" }]
    });

    const results = await page.evaluate(() => {
      const win = app._test.windows[0];
      const row = app._test.tables[win.tableName].rows[0];
      return {
        name: app._test.getDisplayValue(win.tableName, 'name', row),
        email: app._test.getDisplayValue(win.tableName, 'email', row),
      };
    });
    expect(results.name).toBe('ALICE JOHNSON');
    expect(results.email).toBe('alice@example.com');
  });

  test('last-loaded plugin wins for same column', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Plugin A',
      table: '.*',
      columns: [{ match: '^name$', display: "upper(value)" }]
    });
    await loadPluginConfig(page, {
      name: 'Plugin B',
      table: '.*',
      columns: [{ match: '^name$', display: "lower(value)" }]
    });

    const result = await page.evaluate(() => {
      const win = app._test.windows[0];
      return app._test.getDisplayValue(win.tableName, 'name', app._test.tables[win.tableName].rows[0]);
    });
    expect(result).toBe('alice johnson');
  });

  test('unloading later plugin reveals earlier plugin rule', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Plugin A',
      table: '.*',
      columns: [{ match: '^name$', display: "upper(value)" }]
    });
    await loadPluginConfig(page, {
      name: 'Plugin B',
      table: '.*',
      columns: [{ match: '^name$', display: "lower(value)" }]
    });

    await page.evaluate(() => app._test.unloadPlugin(1));

    const result = await page.evaluate(() => {
      const win = app._test.windows[0];
      return app._test.getDisplayValue(win.tableName, 'name', app._test.tables[win.tableName].rows[0]);
    });
    expect(result).toBe('ALICE JOHNSON');
  });

  test('plugin table regex restricts matching', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Only Users',
      table: '^users$',
      columns: [{ match: '^name$', display: "upper(value)" }]
    });

    const result = await page.evaluate(() => {
      const win = app._test.windows[0];
      return app._test.getDisplayValue(win.tableName, 'name', app._test.tables[win.tableName].rows[0]);
    });
    expect(result).toBe('Alice Johnson');
  });

  test('validatePlugin catches missing required fields', async ({ page }) => {
    const errors = await page.evaluate(() => {
      return app._test.validatePlugin({});
    });
    expect(errors.length).toBeGreaterThan(0);
    expect(errors.some(e => e.includes('name'))).toBe(true);
    expect(errors.some(e => e.includes('table'))).toBe(true);
    expect(errors.some(e => e.includes('columns'))).toBe(true);
  });

  test('validatePlugin catches invalid regex', async ({ page }) => {
    const errors = await page.evaluate(() => {
      return app._test.validatePlugin({
        name: 'test',
        table: '[invalid',
        columns: [{ match: 'ok', display: 'value' }]
      });
    });
    expect(errors.some(e => e.includes('regex'))).toBe(true);
  });

  test('compilePlugin skips bad column rules gracefully', async ({ page }) => {
    const result = await page.evaluate(() => {
      const compiled = app._test.compilePlugin({
        name: 'test',
        table: '.*',
        columns: [
          { match: '[bad', display: 'value' },
          { match: '^name$', display: 'upper(value)' }
        ]
      });
      return compiled.columns.length;
    });
    expect(result).toBe(1);
  });

  test('runtime expression error falls back to raw value', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Bad Expr',
      table: '.*',
      columns: [{ match: '^name$', display: "unknownFunc(value)" }]
    });

    const result = await page.evaluate(() => {
      const win = app._test.windows[0];
      return app._test.getDisplayValue(win.tableName, 'name', app._test.tables[win.tableName].rows[0]);
    });
    expect(result).toBe('Alice Johnson');
  });

  test('persistence saves and restores plugins', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Persist Test',
      table: '.*',
      columns: [{ match: '^name$', display: "upper(value)" }]
    });
    await page.evaluate(() => app._test.persistPlugins());

    const stored = await page.evaluate(() => {
      return JSON.parse(localStorage.getItem('csvsql_plugins'));
    });
    expect(stored).toHaveLength(1);
    expect(stored[0].name).toBe('Persist Test');

    await page.evaluate(() => {
      while (app._test.plugins.length) app._test.plugins.pop();
      app._test.rebuildTransformCache();
      app._test.loadPersistedPlugins();
    });

    const result = await page.evaluate(() => {
      const win = app._test.windows[0];
      return app._test.getDisplayValue(win.tableName, 'name', app._test.tables[win.tableName].rows[0]);
    });
    expect(result).toBe('ALICE JOHNSON');
  });

  test('cross-column expression accesses row fields', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Cross Column',
      table: '.*',
      columns: [{ match: '^name$', display: "value + ' <' + row.email + '>'" }]
    });

    const result = await page.evaluate(() => {
      const win = app._test.windows[0];
      return app._test.getDisplayValue(win.tableName, 'name', app._test.tables[win.tableName].rows[0]);
    });
    expect(result).toBe('Alice Johnson <alice@example.com>');
  });

  test('Escape in edit mode re-applies display transform', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Upper Names',
      table: '.*',
      columns: [{ match: '^name$', display: "upper(value)" }]
    });

    // Trigger a re-render so the display transform is applied
    await page.evaluate(() => {
      const win = app._test.windows[0];
      if (win._scrollContainer) win._scrollContainer.dispatchEvent(new Event('scroll'));
    });
    await page.waitForTimeout(200);

    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.keyboard.press('Enter'); // enter edit mode
    await page.waitForTimeout(100);

    // In edit mode, cell should show raw value
    const rawText = await cell.textContent();
    expect(rawText).toBe('Alice Johnson');

    await page.keyboard.press('Escape'); // revert and exit edit mode
    await page.waitForTimeout(100);

    const afterEscape = await cell.textContent();
    expect(afterEscape).toBe('ALICE JOHNSON');
  });

  test('Enter commit in edit mode re-applies display transform', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Upper Names',
      table: '.*',
      columns: [{ match: '^name$', display: "upper(value)" }]
    });

    await page.evaluate(() => {
      const win = app._test.windows[0];
      if (win._scrollContainer) win._scrollContainer.dispatchEvent(new Event('scroll'));
    });
    await page.waitForTimeout(200);

    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.keyboard.press('Enter'); // enter edit mode
    await page.waitForTimeout(100);

    await page.keyboard.press('Enter'); // commit edit
    await page.waitForTimeout(100);

    const afterCommit = await cell.textContent();
    expect(afterCommit).toBe('ALICE JOHNSON');
  });
});
