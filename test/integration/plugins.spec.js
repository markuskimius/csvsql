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
    expect(errors.some(e => e.includes('table') || e.includes('tables') || e.includes('links'))).toBe(true);
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
      return compiled.tables[0].columns.length;
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

  test('loadPluginFile loads a plugin from a File object', async ({ page }) => {
    const result = await page.evaluate(async () => {
      const json = JSON.stringify({
        name: 'Drop Test',
        table: '.*',
        columns: [{ match: '^name$', display: "upper(value)" }]
      });
      const file = new File([json], 'drop-test.json', { type: 'application/json' });
      await app._test.loadPluginFile(file);
      return {
        count: app._test.plugins.length,
        name: app._test.plugins[0]?.name,
        filename: app._test.plugins[0]?._filename,
        transformed: app._test.getDisplayValue(
          app._test.windows[0].tableName, 'name',
          app._test.tables[app._test.windows[0].tableName].rows[0]
        ),
      };
    });
    expect(result.count).toBe(1);
    expect(result.name).toBe('Drop Test');
    expect(result.filename).toBe('drop-test.json');
    expect(result.transformed).toBe('ALICE JOHNSON');
  });

  test('dropping a .json file loads it as a plugin, not a data file', async ({ page }) => {
    const result = await page.evaluate(async () => {
      const json = JSON.stringify({
        name: 'Drop Plugin',
        table: '.*',
        columns: [{ match: '^name$', display: "lower(value)" }]
      });
      const file = new File([json], 'test-plugin.json', { type: 'application/json' });
      const dt = new DataTransfer();
      dt.items.add(file);
      const evt = new DragEvent('drop', { bubbles: true, dataTransfer: dt });
      document.dispatchEvent(evt);
      await new Promise(r => setTimeout(r, 200));
      const windowCount = app._test.windows.length;
      const pluginCount = app._test.plugins.length;
      const hasTransform = app._test.hasDisplayTransform(
        app._test.windows[0].tableName, 'name'
      );
      return { windowCount, pluginCount, hasTransform };
    });
    // Should still have just the original sample1 window (no new data window)
    expect(result.windowCount).toBe(1);
    // Plugin was loaded
    expect(result.pluginCount).toBe(1);
    expect(result.hasTransform).toBe(true);
  });

  test('plugin toggle appears in status bar when plugins match', async ({ page }) => {
    let toggle = await page.$('.status-plugin-toggle');
    expect(toggle).toBeNull();

    await loadPluginConfig(page, {
      name: 'Toggle Test', table: '.*',
      columns: [{ match: '^name$', display: 'upper(value)' }]
    });
    await page.evaluate(() => app._test.rerenderAllWindows());
    await page.waitForTimeout(100);

    toggle = await page.$('.status-plugin-toggle');
    expect(toggle).not.toBeNull();
    const text = await toggle.textContent();
    expect(text).toContain('Plugins on');
  });

  test('plugin toggle does not appear when no plugins match', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'No Match', table: '^nonexistent$',
      columns: [{ match: '^name$', display: 'upper(value)' }]
    });
    await page.evaluate(() => app._test.rerenderAllWindows());
    await page.waitForTimeout(100);

    const toggle = await page.$('.status-plugin-toggle');
    expect(toggle).toBeNull();
  });

  test('status bar toggle disables all transforms, icons become disabled', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Toggle Disable', table: '.*',
      columns: [{ match: '^name$', display: 'upper(value)' }]
    });
    await page.evaluate(() => app._test.rerenderAllWindows());
    await page.waitForTimeout(100);

    const before = await getCellText(page, 0, 0);
    expect(before).toBe('ALICE JOHNSON');

    let icon = await page.$('.col-transform-icon');
    expect(icon).not.toBeNull();
    expect(await icon.evaluate(el => el.classList.contains('disabled'))).toBe(false);

    await page.click('.status-plugin-toggle');
    await page.waitForTimeout(100);

    const after = await getCellText(page, 0, 0);
    expect(after).toBe('Alice Johnson');

    const toggleText = await page.$eval('.status-plugin-toggle', el => el.textContent);
    expect(toggleText).toContain('Plugins off');

    icon = await page.$('.col-transform-icon');
    expect(icon).not.toBeNull();
    expect(await icon.evaluate(el => el.classList.contains('disabled'))).toBe(true);
  });

  test('status bar toggle re-enables all transforms', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Toggle Re-enable', table: '.*',
      columns: [{ match: '^name$', display: 'upper(value)' }]
    });
    await page.evaluate(() => app._test.rerenderAllWindows());
    await page.waitForTimeout(100);

    await page.click('.status-plugin-toggle');
    await page.waitForTimeout(100);
    await page.click('.status-plugin-toggle');
    await page.waitForTimeout(100);

    const cell = await getCellText(page, 0, 0);
    expect(cell).toBe('ALICE JOHNSON');

    const toggleText = await page.$eval('.status-plugin-toggle', el => el.textContent);
    expect(toggleText).toContain('Plugins on');

    const icon = await page.$('.col-transform-icon');
    expect(await icon.evaluate(el => el.classList.contains('disabled'))).toBe(false);
  });

  test('clicking column plug icon disables only that column', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Multi Col', table: '.*',
      columns: [
        { match: '^name$', display: 'upper(value)' },
        { match: '^email$', display: 'upper(value)' }
      ]
    });
    await page.evaluate(() => app._test.rerenderAllWindows());
    await page.waitForTimeout(100);

    const nameBefore = await getCellText(page, 0, 0);
    expect(nameBefore).toBe('ALICE JOHNSON');

    const icons = await page.$$('.col-transform-icon');
    expect(icons.length).toBe(2);
    await icons[0].click();
    await page.waitForTimeout(100);

    const nameAfter = await getCellText(page, 0, 0);
    expect(nameAfter).toBe('Alice Johnson');

    const emailIdx = await page.evaluate(() => app._test.windows[0]._columns.indexOf('email'));
    const emailCell = await getCellText(page, 0, emailIdx);
    expect(emailCell).toBe('ALICE@EXAMPLE.COM');

    const toggleText = await page.$eval('.status-plugin-toggle', el => el.textContent);
    expect(toggleText).toContain('Plugins partial');
  });

  test('clicking disabled column plug icon re-enables that column', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Re-enable Col', table: '.*',
      columns: [{ match: '^name$', display: 'upper(value)' }]
    });
    await page.evaluate(() => app._test.rerenderAllWindows());
    await page.waitForTimeout(100);

    const icon = await page.$('.col-transform-icon');
    await icon.click();
    await page.waitForTimeout(100);
    expect(await getCellText(page, 0, 0)).toBe('Alice Johnson');

    const disabledIcon = await page.$('.col-transform-icon');
    await disabledIcon.click();
    await page.waitForTimeout(100);
    expect(await getCellText(page, 0, 0)).toBe('ALICE JOHNSON');

    const reenabledIcon = await page.$('.col-transform-icon');
    expect(await reenabledIcon.evaluate(el => el.classList.contains('disabled'))).toBe(false);
  });

  test('column icon only appears on transformed columns', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Selective', table: '.*',
      columns: [{ match: '^name$', display: 'upper(value)' }]
    });
    await page.evaluate(() => app._test.rerenderAllWindows());
    await page.waitForTimeout(100);

    const transformed = await page.$$eval('.col-transform-icon', els => els.map(el => {
      const th = el.closest('th');
      const colName = th && th.querySelector('.col-name');
      return colName ? colName.textContent.trim() : '';
    }));
    expect(transformed).toContain('name');
    expect(transformed).not.toContain('age');
    expect(transformed).not.toContain('city');
  });

  test('disabled transform state survives table rebuild from sort', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Sort Survive', table: '.*',
      columns: [{ match: '^name$', display: 'upper(value)' }]
    });
    await page.evaluate(() => app._test.rerenderAllWindows());
    await page.waitForTimeout(100);

    await page.click('.col-transform-icon');
    await page.waitForTimeout(100);

    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const cell = await getCellText(page, 0, 0);
    expect(cell).toBe('Alice Johnson');

    const icon = await page.$('.col-transform-icon');
    expect(await icon.evaluate(el => el.classList.contains('disabled'))).toBe(true);
  });

  test('autofilter dropdown shows display-formatted values when transform active', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Upper Names', table: '.*',
      columns: [{ match: '^name$', display: 'upper(value)' }]
    });
    await page.evaluate(() => app._test.rerenderAllWindows());
    await page.waitForTimeout(100);

    const filterBtn = await page.$('th .col-filter-btn');
    await filterBtn.click();
    await page.waitForTimeout(200);

    const items = await page.$$eval('.autofilter-item span', els => els.map(el => el.textContent));
    expect(items.length).toBeGreaterThan(0);
    expect(items.every(t => t === t.toUpperCase())).toBe(true);
    expect(items).toContain('ALICE JOHNSON');
  });

  test('autofilter dropdown shows raw values when column transform disabled', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Upper Names', table: '.*',
      columns: [{ match: '^name$', display: 'upper(value)' }]
    });
    await page.evaluate(() => app._test.rerenderAllWindows());
    await page.waitForTimeout(100);

    await page.click('.col-transform-icon');
    await page.waitForTimeout(100);

    const filterBtn = await page.$('th .col-filter-btn');
    await filterBtn.click();
    await page.waitForTimeout(200);

    const items = await page.$$eval('.autofilter-item span', els => els.map(el => el.textContent));
    expect(items).toContain('Alice Johnson');
  });

  test('autofilter search matches display-formatted values', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Upper Names', table: '.*',
      columns: [{ match: '^name$', display: 'upper(value)' }]
    });
    await page.evaluate(() => app._test.rerenderAllWindows());
    await page.waitForTimeout(100);

    const filterBtn = await page.$('th .col-filter-btn');
    await filterBtn.click();
    await page.waitForTimeout(200);

    const searchInput = await page.$('.autofilter-search');
    await searchInput.fill('ALICE');
    await page.waitForTimeout(200);

    const items = await page.$$eval('.autofilter-item span', els => els.map(el => el.textContent));
    expect(items).toContain('ALICE JOHNSON');
    expect(items.length).toBe(1);
  });

  test('edit mode shows raw value when per-column transform is active', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Upper Names', table: '.*',
      columns: [{ match: '^name$', display: 'upper(value)' }]
    });
    await page.evaluate(() => app._test.rerenderAllWindows());
    await page.waitForTimeout(100);

    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    expect(await cell.textContent()).toBe('ALICE JOHNSON');

    await cell.click();
    await page.keyboard.press('Enter');
    await page.waitForTimeout(100);
    expect(await cell.textContent()).toBe('Alice Johnson');

    await page.keyboard.press('Escape');
    await page.waitForTimeout(100);
    expect(await cell.textContent()).toBe('ALICE JOHNSON');
  });

  test('edit mode with per-column transform disabled shows raw value throughout', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Upper Names', table: '.*',
      columns: [{ match: '^name$', display: 'upper(value)' }]
    });
    await page.evaluate(() => app._test.rerenderAllWindows());
    await page.waitForTimeout(100);

    await page.click('.col-transform-icon');
    await page.waitForTimeout(100);

    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    expect(await cell.textContent()).toBe('Alice Johnson');

    await cell.click();
    await page.keyboard.press('Enter');
    await page.waitForTimeout(100);
    expect(await cell.textContent()).toBe('Alice Johnson');

    await page.keyboard.press('Escape');
    await page.waitForTimeout(100);
    expect(await cell.textContent()).toBe('Alice Johnson');
  });

  test('status bar shows partial when some columns disabled', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Multi', table: '.*',
      columns: [
        { match: '^name$', display: 'upper(value)' },
        { match: '^email$', display: 'upper(value)' }
      ]
    });
    await page.evaluate(() => app._test.rerenderAllWindows());
    await page.waitForTimeout(100);

    let toggleText = await page.$eval('.status-plugin-toggle', el => el.textContent);
    expect(toggleText).toContain('Plugins on');

    const icons = await page.$$('.col-transform-icon');
    await icons[0].click();
    await page.waitForTimeout(100);

    toggleText = await page.$eval('.status-plugin-toggle', el => el.textContent);
    expect(toggleText).toContain('Plugins partial');

    await page.click('.status-plugin-toggle');
    await page.waitForTimeout(100);

    toggleText = await page.$eval('.status-plugin-toggle', el => el.textContent);
    expect(toggleText).toContain('Plugins on');
    expect(await getCellText(page, 0, 0)).toBe('ALICE JOHNSON');
  });

  test('unloading plugin removes icons and toggle', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Removable', table: '.*',
      columns: [{ match: '^name$', display: 'upper(value)' }]
    });
    await page.evaluate(() => app._test.rerenderAllWindows());
    await page.waitForTimeout(100);

    expect(await page.$('.col-transform-icon')).not.toBeNull();
    expect(await page.$('.status-plugin-toggle')).not.toBeNull();

    await page.evaluate(() => app._test.unloadPlugin(0));
    await page.waitForTimeout(100);

    expect(await page.$('.col-transform-icon')).toBeNull();
    expect(await page.$('.status-plugin-toggle')).toBeNull();
  });

  test('autofilter still filters correctly with display-formatted dropdown', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Upper Names', table: '.*',
      columns: [{ match: '^name$', display: 'upper(value)' }]
    });
    await page.evaluate(() => app._test.rerenderAllWindows());
    await page.waitForTimeout(100);

    const totalRows = await page.$$eval(
      '.subwindow table tbody tr:not(.virtual-pad)', rows => rows.length
    );

    const filterBtn = await page.$('th .col-filter-btn');
    await filterBtn.click();
    await page.waitForTimeout(200);

    // Uncheck Select All, then check only the first value item
    const selectAllCb = await page.$('.autofilter-select-all input[type="checkbox"]');
    await selectAllCb.click();
    await page.waitForTimeout(50);
    const firstItemCb = await page.$('.autofilter-item input[type="checkbox"]');
    await firstItemCb.click();
    await page.waitForTimeout(50);

    const applyBtn = await page.$('.autofilter-apply');
    await applyBtn.click();
    await page.waitForTimeout(200);

    const filteredRows = await page.$$eval(
      '.subwindow table tbody tr:not(.virtual-pad)', rows => rows.length
    );
    expect(filteredRows).toBeLessThan(totalRows);
    expect(filteredRows).toBeGreaterThan(0);
  });

  test('validatePlugin accepts optional metadata fields', async ({ page }) => {
    const result = await page.evaluate(() => {
      const config = {
        name: 'Meta Test', version: '2.0.0', author: 'Test Author',
        created: '2026-01-01', description: 'A test plugin',
        table: '.*', columns: [{ match: '^name$', display: 'value' }]
      };
      return app._test.validatePlugin(config);
    });
    expect(result).toEqual([]);
  });

  test('validatePlugin rejects non-string metadata fields', async ({ page }) => {
    const result = await page.evaluate(() => {
      const config = {
        name: 'Bad Meta', table: '.*',
        columns: [{ match: '^name$', display: 'value' }],
        version: 123, author: true, created: 42
      };
      return app._test.validatePlugin(config);
    });
    expect(result).toContain('"version" must be a string');
    expect(result).toContain('"author" must be a string');
    expect(result).toContain('"created" must be a string');
  });

  test('metadata fields persist through save/load cycle', async ({ page }) => {
    const result = await page.evaluate(async () => {
      const json = JSON.stringify({
        name: 'Persist Meta', version: '1.2.3', author: 'Alice',
        created: '2026-06-01', description: 'Persists metadata',
        table: '.*', columns: [{ match: '^name$', display: 'upper(value)' }]
      });
      const file = new File([json], 'meta.json', { type: 'application/json' });
      await app._test.loadPluginFile(file);
      const stored = JSON.parse(localStorage.getItem('csvsql_plugins'));
      return stored[0];
    });
    expect(result.version).toBe('1.2.3');
    expect(result.author).toBe('Alice');
    expect(result.created).toBe('2026-06-01');
    expect(result.description).toBe('Persists metadata');
  });

  test('showToast creates and auto-removes toast element', async ({ page }) => {
    await page.evaluate(() => app._test.showToast('Test message'));
    const toast = await page.$('.toast.toast-success');
    expect(toast).not.toBeNull();
    expect(await toast.textContent()).toBe('Test message');
    await page.waitForTimeout(3500);
    const remaining = await page.$('.toast');
    expect(remaining).toBeNull();
  });

  test('showToast error type applies error class', async ({ page }) => {
    await page.evaluate(() => app._test.showToast('Error msg', 'error'));
    const toast = await page.$('.toast.toast-error');
    expect(toast).not.toBeNull();
    expect(await toast.textContent()).toBe('Error msg');
  });

  test('multiple toasts stack vertically', async ({ page }) => {
    await page.evaluate(() => {
      app._test.showToast('First');
      app._test.showToast('Second');
    });
    const toasts = await page.$$('.toast');
    expect(toasts.length).toBe(2);
    const bottom0 = await toasts[0].evaluate(el => parseFloat(el.style.bottom));
    const bottom1 = await toasts[1].evaluate(el => parseFloat(el.style.bottom));
    expect(bottom1).toBeGreaterThan(bottom0);
  });

  test('loadPluginFile shows success toast instead of alert', async ({ page }) => {
    const result = await page.evaluate(async () => {
      const json = JSON.stringify({
        name: 'Toast Test', table: '.*',
        columns: [{ match: '^name$', display: 'upper(value)' }]
      });
      const file = new File([json], 'toast.json', { type: 'application/json' });
      await app._test.loadPluginFile(file);
      const toast = document.querySelector('.toast.toast-success');
      return toast ? toast.textContent : null;
    });
    expect(result).toContain('Toast Test');
    expect(result).toContain('loaded');
  });

  test('loadPluginFile shows error toast for invalid plugin', async ({ page }) => {
    const result = await page.evaluate(async () => {
      const json = JSON.stringify({ name: '', table: '.*', columns: [] });
      const file = new File([json], 'bad.json', { type: 'application/json' });
      await app._test.loadPluginFile(file);
      const toast = document.querySelector('.toast.toast-error');
      return toast ? toast.textContent : null;
    });
    expect(result).not.toBeNull();
    expect(result).toContain('error');
  });

  test('unloadPlugin shows success toast', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Unload Toast', table: '.*',
      columns: [{ match: '^name$', display: 'upper(value)' }]
    });
    const result = await page.evaluate(() => {
      app._test.unloadPlugin(0);
      const toast = document.querySelector('.toast.toast-success');
      return toast ? toast.textContent : null;
    });
    expect(result).toContain('Unload Toast');
    expect(result).toContain('unloaded');
  });

  test('plugin menu entry has unload button and clickable name', async ({ page }) => {
    await page.evaluate(async () => {
      const json = JSON.stringify({
        name: 'Menu Entry', table: '.*',
        columns: [{ match: '^name$', display: 'upper(value)' }]
      });
      const file = new File([json], 'menu.json', { type: 'application/json' });
      await app._test.loadPluginFile(file);
    });
    const entry = await page.$('.plugin-entry');
    expect(entry).not.toBeNull();
    const unload = await page.$('.plugin-unload');
    expect(unload).not.toBeNull();
    expect(await unload.textContent()).toBe('✕');
    const name = await page.$('.plugin-name');
    expect(name).not.toBeNull();
    expect(await name.textContent()).toBe('Menu Entry');
  });

  test('clicking plugin name opens about window with metadata', async ({ page }) => {
    await page.evaluate(async () => {
      const json = JSON.stringify({
        name: 'About Test', version: '3.0.0', author: 'Bob',
        created: '2026-03-15', description: 'Test description',
        table: '.*', columns: [{ match: '^name$', display: 'upper(value)' }]
      });
      const file = new File([json], 'about.json', { type: 'application/json' });
      await app._test.loadPluginFile(file);
    });
    await page.click('#menu-plugins .menu-label');
    await page.waitForTimeout(100);
    await page.click('.plugin-name');
    await page.waitForTimeout(100);

    const aboutWin = page.locator('.subwindow').filter({ hasText: 'Plugin: About Test' });
    await expect(aboutWin).toBeVisible();
    const text = await aboutWin.textContent();
    expect(text).toContain('3.0.0');
    expect(text).toContain('Bob');
    expect(text).toContain('2026-03-15');
    expect(text).toContain('Test description');
    expect(text).toContain('Unload Plugin');
  });

  test('about window close button dismisses it', async ({ page }) => {
    await page.evaluate(async () => {
      const json = JSON.stringify({
        name: 'Close Test', table: '.*',
        columns: [{ match: '^name$', display: 'upper(value)' }]
      });
      const file = new File([json], 'close.json', { type: 'application/json' });
      await app._test.loadPluginFile(file);
    });
    await page.click('#menu-plugins .menu-label');
    await page.waitForTimeout(100);
    await page.click('.plugin-name');
    await page.waitForTimeout(100);

    const aboutWin = page.locator('.subwindow').filter({ hasText: 'Plugin: Close Test' });
    await expect(aboutWin).toBeVisible();

    await aboutWin.locator('.btn-close').click();
    await page.waitForTimeout(100);
    await expect(aboutWin).not.toBeAttached();
  });

  test('about window unload button unloads plugin and closes', async ({ page }) => {
    await page.evaluate(async () => {
      const json = JSON.stringify({
        name: 'Unload Via Dialog', table: '.*',
        columns: [{ match: '^name$', display: 'upper(value)' }]
      });
      const file = new File([json], 'unload-dialog.json', { type: 'application/json' });
      await app._test.loadPluginFile(file);
    });
    await page.click('#menu-plugins .menu-label');
    await page.waitForTimeout(100);
    await page.click('.plugin-name');
    await page.waitForTimeout(100);

    const aboutWin = page.locator('.subwindow').filter({ hasText: 'Plugin: Unload Via Dialog' });
    await aboutWin.locator('.plugin-about-unload').click();
    await page.waitForTimeout(100);

    await expect(aboutWin).not.toBeAttached();
    const pluginCount = await page.evaluate(() => app._test.plugins.length);
    expect(pluginCount).toBe(0);
  });

  test('about window omits metadata fields when absent', async ({ page }) => {
    await page.evaluate(async () => {
      const json = JSON.stringify({
        name: 'Minimal Plugin', table: '.*',
        columns: [{ match: '^name$', display: 'upper(value)' }]
      });
      const file = new File([json], 'minimal.json', { type: 'application/json' });
      await app._test.loadPluginFile(file);
    });
    await page.click('#menu-plugins .menu-label');
    await page.waitForTimeout(100);
    await page.click('.plugin-name');
    await page.waitForTimeout(100);

    const aboutWin = page.locator('.subwindow').filter({ hasText: 'Plugin: Minimal Plugin' });
    const text = await aboutWin.textContent();
    expect(text).not.toContain('Version:');
    expect(text).not.toContain('Author:');
    expect(text).not.toContain('Created:');
  });

  test('about window shows column rules', async ({ page }) => {
    await page.evaluate(async () => {
      const json = JSON.stringify({
        name: 'Rules Test', table: 'sample.*',
        columns: [
          { match: '^name$', display: 'upper(value)' },
          { match: '^age$', display: 'value + " years"' }
        ]
      });
      const file = new File([json], 'rules.json', { type: 'application/json' });
      await app._test.loadPluginFile(file);
    });
    await page.click('#menu-plugins .menu-label');
    await page.waitForTimeout(100);
    await page.click('.plugin-name');
    await page.waitForTimeout(100);

    const aboutWin = page.locator('.subwindow').filter({ hasText: 'Plugin: Rules Test' });
    const text = await aboutWin.textContent();
    expect(text).toContain('sample.*');
    expect(text).toContain('^name$');
    expect(text).toContain('upper(value)');
    expect(text).toContain('^age$');
    expect(text).toContain('value + " years"');
  });

  test('menu unload button still works directly', async ({ page }) => {
    await page.evaluate(async () => {
      const json = JSON.stringify({
        name: 'Direct Unload', table: '.*',
        columns: [{ match: '^name$', display: 'upper(value)' }]
      });
      const file = new File([json], 'direct.json', { type: 'application/json' });
      await app._test.loadPluginFile(file);
    });
    await page.click('#menu-plugins .menu-label');
    await page.waitForTimeout(100);
    await page.click('.plugin-unload');
    await page.waitForTimeout(100);

    const pluginCount = await page.evaluate(() => app._test.plugins.length);
    expect(pluginCount).toBe(0);
  });

  test('multi-table plugin applies rules to different tables', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Multi Table',
      tables: [
        { table: '^sample1$', columns: [{ match: '^name$', display: 'upper(value)' }] },
        { table: '^other$', columns: [{ match: '^name$', display: 'lower(value)' }] }
      ]
    });

    const result = await page.evaluate(() => {
      const win = app._test.windows[0];
      const row = app._test.tables[win.tableName].rows[0];
      return app._test.getDisplayValue(win.tableName, 'name', row);
    });
    expect(result).toBe('ALICE JOHNSON');
  });

  test('old format plugin still works (backward compat)', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Old Format',
      table: '.*',
      columns: [{ match: '^name$', display: 'upper(value)' }]
    });

    const result = await page.evaluate(() => {
      const win = app._test.windows[0];
      return app._test.getDisplayValue(win.tableName, 'name', app._test.tables[win.tableName].rows[0]);
    });
    expect(result).toBe('ALICE JOHNSON');
  });

  test('validatePlugin accepts tables array format', async ({ page }) => {
    const errors = await page.evaluate(() => {
      return app._test.validatePlugin({
        name: 'Multi',
        tables: [
          { table: '.*', columns: [{ match: '^name$', display: 'value' }] }
        ]
      });
    });
    expect(errors).toEqual([]);
  });

  test('validatePlugin accepts links-only plugin', async ({ page }) => {
    const errors = await page.evaluate(() => {
      return app._test.validatePlugin({
        name: 'Links Only',
        links: [
          { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
        ]
      });
    });
    expect(errors).toEqual([]);
  });

  test('validatePlugin catches invalid link fields', async ({ page }) => {
    const errors = await page.evaluate(() => {
      return app._test.validatePlugin({
        name: 'Bad Link',
        links: [
          { source: { table: '[bad', column: 'id' }, target: { table: 'ok', column: 'id' } }
        ]
      });
    });
    expect(errors.some(e => e.includes('regex'))).toBe(true);
  });

  test('compilePlugin compiles links', async ({ page }) => {
    const result = await page.evaluate(() => {
      const compiled = app._test.compilePlugin({
        name: 'Link Test',
        links: [
          { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
        ]
      });
      return { tables: compiled.tables.length, links: compiled.links.length };
    });
    expect(result.tables).toBe(0);
    expect(result.links).toBe(1);
  });

  test('about window shows links section', async ({ page }) => {
    await page.evaluate(async () => {
      const json = JSON.stringify({
        name: 'Link About',
        tables: [{ table: '.*', columns: [{ match: '^name$', display: 'upper(value)' }] }],
        links: [
          { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
        ]
      });
      const file = new File([json], 'link-about.json', { type: 'application/json' });
      await app._test.loadPluginFile(file);
      app._test.showPluginAbout(app._test.plugins[0], 0);
    });
    await page.waitForTimeout(100);

    const aboutWin = page.locator('.subwindow').filter({ hasText: 'Plugin: Link About' });
    const text = await aboutWin.textContent();
    expect(text).toContain('Links:');
    expect(text).toContain('orders.customer_id');
    expect(text).toContain('customers.id');
  });

  test('validatePlugin accepts combo plugin with tables and links', async ({ page }) => {
    const errors = await page.evaluate(() => {
      return app._test.validatePlugin({
        name: 'Combo',
        tables: [{ table: '.*', columns: [{ match: '^name$', display: 'value' }] }],
        links: [
          { source: { table: 'a', column: 'id' }, target: { table: 'b', column: 'a_id' } }
        ]
      });
    });
    expect(errors).toEqual([]);
  });

  test('validatePlugin catches tables entry missing table field', async ({ page }) => {
    const errors = await page.evaluate(() => {
      return app._test.validatePlugin({
        name: 'Bad Entry',
        tables: [{ columns: [{ match: '^name$', display: 'value' }] }]
      });
    });
    expect(errors.length).toBeGreaterThan(0);
  });

  test('validatePlugin catches link entry missing source or target', async ({ page }) => {
    const errors = await page.evaluate(() => {
      return app._test.validatePlugin({
        name: 'Bad Link',
        links: [{ source: { table: 'a', column: 'id' } }]
      });
    });
    expect(errors.length).toBeGreaterThan(0);
  });

  test('multi-table plugin second table entry does not match first table', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Multi Table Isolation',
      tables: [
        { table: '^sample1$', columns: [{ match: '^name$', display: 'upper(value)' }] },
        { table: '^other$', columns: [{ match: '^name$', display: 'lower(value)' }] }
      ]
    });

    const result = await page.evaluate(() => {
      const win = app._test.windows[0];
      return {
        tableName: win.tableName,
        display: app._test.getDisplayValue(win.tableName, 'name', app._test.tables[win.tableName].rows[0]),
      };
    });
    expect(result.tableName).toBe('sample1');
    expect(result.display).toBe('ALICE JOHNSON');
  });

  test('about window shows multi-table entries', async ({ page }) => {
    await page.evaluate(async () => {
      const json = JSON.stringify({
        name: 'Multi About',
        tables: [
          { table: 'orders', columns: [{ match: '^total$', display: 'value' }] },
          { table: 'customers', columns: [{ match: '^name$', display: 'upper(value)' }] }
        ]
      });
      const file = new File([json], 'multi-about.json', { type: 'application/json' });
      await app._test.loadPluginFile(file);
      app._test.showPluginAbout(app._test.plugins[0], 0);
    });
    await page.waitForTimeout(100);

    const aboutWin = page.locator('.subwindow').filter({ hasText: 'Plugin: Multi About' });
    const text = await aboutWin.textContent();
    expect(text).toContain('orders');
    expect(text).toContain('customers');
    expect(text).toContain('^total$');
    expect(text).toContain('^name$');
  });
});

test.describe('Cross-table linking', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, 'customers.csv');
    await waitForWindow(page, 'customers');
    await uploadFile(page, 'orders.csv');
    await waitForWindow(page, 'orders');
  });

  async function loadLinkPlugin(page, config) {
    await page.evaluate((cfg) => {
      const errors = app._test.validatePlugin(cfg);
      if (errors.length) throw new Error(errors.join(', '));
      const compiled = app._test.compilePlugin(cfg);
      cfg._compiled = compiled;
      app._test.plugins.push(cfg);
      app._test.rebuildTransformCache();
    }, config);
  }

  test('selecting a row in source filters target table', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Link Test',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    const result = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      const ordersTable = app._test.tables['orders'];
      const row = ordersTable.rows[0]; // order 101, customer_id=1

      ordersWin.anchorCell = { rownum: row._rownum, col: 'order_id' };
      ordersWin.selectedCells = new Set();
      for (const col of ordersTable.columns) {
        ordersWin.selectedCells.add(`${row._rownum}:${col}`);
      }

      app._test.applyLinkFilters(ordersWin);

      return {
        linkFilterKeys: Object.keys(custWin.linkFilters),
        filterValues: custWin.linkFilters['id'] ? [...custWin.linkFilters['id']] : [],
      };
    });

    expect(result.linkFilterKeys).toEqual(['id']);
    expect(result.filterValues).toEqual(['1']);
  });

  test('clearing selection clears link filters on target', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Link Test',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const ordersTable = app._test.tables['orders'];
      const row = ordersTable.rows[0];

      ordersWin.anchorCell = { rownum: row._rownum, col: 'order_id' };
      ordersWin.selectedCells = new Set();
      for (const col of ordersTable.columns) {
        ordersWin.selectedCells.add(`${row._rownum}:${col}`);
      }
      app._test.applyLinkFilters(ordersWin);
    });

    const afterClear = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      ordersWin.selectedCells = new Set();
      app._test.clearLinkFilters(ordersWin);
      return Object.keys(custWin.linkFilters).length;
    });

    expect(afterClear).toBe(0);
  });

  test('multi-row selection collects all source values', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Link Test',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    const result = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      const ordersTable = app._test.tables['orders'];

      // Select rows 0 and 2 (customer_id 1 and 2)
      ordersWin.selectedCells = new Set();
      for (const row of [ordersTable.rows[0], ordersTable.rows[2]]) {
        for (const col of ordersTable.columns) {
          ordersWin.selectedCells.add(`${row._rownum}:${col}`);
        }
      }
      ordersWin.anchorCell = { rownum: ordersTable.rows[0]._rownum, col: 'order_id' };

      app._test.applyLinkFilters(ordersWin);

      return [...custWin.linkFilters['id']].sort();
    });

    expect(result).toEqual(['1', '2']);
  });

  test('source table is excluded from own link targets', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Self Link',
      links: [
        { source: { table: '.*', column: 'id' }, target: { table: '.*', column: 'id' } }
      ]
    });

    const result = await page.evaluate(() => {
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      const custTable = app._test.tables['customers'];
      const row = custTable.rows[0];

      custWin.anchorCell = { rownum: row._rownum, col: 'id' };
      custWin.selectedCells = new Set();
      for (const col of custTable.columns) {
        custWin.selectedCells.add(`${row._rownum}:${col}`);
      }

      app._test.applyLinkFilters(custWin);

      return Object.keys(custWin.linkFilters).length;
    });

    expect(result).toBe(0);
  });

  test('bidirectional links work both ways', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Bidirectional',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } },
        { source: { table: 'customers', column: 'id' }, target: { table: 'orders', column: 'customer_id' } }
      ]
    });

    const forward = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      const ordersTable = app._test.tables['orders'];
      const row = ordersTable.rows[0]; // customer_id=1

      ordersWin.selectedCells = new Set();
      for (const col of ordersTable.columns) {
        ordersWin.selectedCells.add(`${row._rownum}:${col}`);
      }
      ordersWin.anchorCell = { rownum: row._rownum, col: 'order_id' };
      app._test.applyLinkFilters(ordersWin);

      return [...custWin.linkFilters['id']];
    });
    expect(forward).toEqual(['1']);

    const reverse = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      const custTable = app._test.tables['customers'];

      // Clear previous link filters
      ordersWin.linkFilters = {};
      custWin.linkFilters = {};

      const row = custTable.rows[1]; // id=2

      custWin.selectedCells = new Set();
      for (const col of custTable.columns) {
        custWin.selectedCells.add(`${row._rownum}:${col}`);
      }
      custWin.anchorCell = { rownum: row._rownum, col: 'id' };
      app._test.applyLinkFilters(custWin);

      return [...ordersWin.linkFilters['customer_id']];
    });
    expect(reverse).toEqual(['2']);
  });

  test('link filters apply in buildTableHTML', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Filter Test',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    const result = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      const ordersTable = app._test.tables['orders'];
      const row = ordersTable.rows[0]; // customer_id=1

      ordersWin.selectedCells = new Set();
      for (const col of ordersTable.columns) {
        ordersWin.selectedCells.add(`${row._rownum}:${col}`);
      }
      ordersWin.anchorCell = { rownum: row._rownum, col: 'order_id' };
      app._test.applyLinkFilters(ordersWin);

      // custWin should now have link filter; check display rows after rebuild
      return {
        displayRowCount: custWin._displayRows ? custWin._displayRows.length : -1,
        totalRows: app._test.tables['customers'].rows.length,
      };
    });

    expect(result.totalRows).toBe(3);
    expect(result.displayRowCount).toBe(1);
  });

  test('col-linked class appears on link-filtered columns', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'CSS Test',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const ordersTable = app._test.tables['orders'];
      const row = ordersTable.rows[0];
      ordersWin.selectedCells = new Set();
      for (const col of ordersTable.columns) {
        ordersWin.selectedCells.add(`${row._rownum}:${col}`);
      }
      ordersWin.anchorCell = { rownum: row._rownum, col: 'order_id' };
      app._test.applyLinkFilters(ordersWin);
    });
    await page.waitForTimeout(200);

    const custWindow = page.locator('.subwindow').filter({ hasText: 'customers' });
    const linkedHeaders = await custWindow.locator('th.col-linked').count();
    expect(linkedHeaders).toBeGreaterThan(0);
  });

  test('status bar shows linked indicator', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Status Test',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const ordersTable = app._test.tables['orders'];
      const row = ordersTable.rows[0];
      ordersWin.selectedCells = new Set();
      for (const col of ordersTable.columns) {
        ordersWin.selectedCells.add(`${row._rownum}:${col}`);
      }
      ordersWin.anchorCell = { rownum: row._rownum, col: 'order_id' };
      app._test.applyLinkFilters(ordersWin);
    });
    await page.waitForTimeout(200);

    const custWindow = page.locator('.subwindow').filter({ hasText: 'customers' });
    const linkLabel = custWindow.locator('.status-link-filter');
    await expect(linkLabel).toBeVisible();
    expect(await linkLabel.textContent()).toContain('Linked');
  });

  test('unloading link plugin clears link filters', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Unload Link',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const ordersTable = app._test.tables['orders'];
      const row = ordersTable.rows[0];
      ordersWin.selectedCells = new Set();
      for (const col of ordersTable.columns) {
        ordersWin.selectedCells.add(`${row._rownum}:${col}`);
      }
      ordersWin.anchorCell = { rownum: row._rownum, col: 'order_id' };
      app._test.applyLinkFilters(ordersWin);
    });

    const afterUnload = await page.evaluate(() => {
      app._test.unloadPlugin(0);
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      return Object.keys(custWin.linkFilters).length;
    });
    expect(afterUnload).toBe(0);
  });

  test('link filters coexist with manual column autofilters', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Coexist Test',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    const result = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      const ordersTable = app._test.tables['orders'];

      // Set a manual column autofilter on customers
      custWin.columnFilters = { 'name': new Set(['Alice', 'Bob']) };

      // Apply link filter from orders
      const row = ordersTable.rows[0]; // customer_id=1
      ordersWin.selectedCells = new Set();
      for (const col of ordersTable.columns) {
        ordersWin.selectedCells.add(`${row._rownum}:${col}`);
      }
      ordersWin.anchorCell = { rownum: row._rownum, col: 'order_id' };
      app._test.applyLinkFilters(ordersWin);

      return {
        hasLinkFilter: Object.keys(custWin.linkFilters).length > 0,
        hasColumnFilter: Object.keys(custWin.columnFilters).length > 0,
      };
    });

    expect(result.hasLinkFilter).toBe(true);
    expect(result.hasColumnFilter).toBe(true);
  });

  test('empty selection via applyLinkFilters clears target filters', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Empty Sel Test',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    const result = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      const ordersTable = app._test.tables['orders'];

      // First select a row to create link filter
      const row = ordersTable.rows[0];
      ordersWin.selectedCells = new Set();
      for (const col of ordersTable.columns) {
        ordersWin.selectedCells.add(`${row._rownum}:${col}`);
      }
      ordersWin.anchorCell = { rownum: row._rownum, col: 'order_id' };
      app._test.applyLinkFilters(ordersWin);
      const hadFilter = Object.keys(custWin.linkFilters).length > 0;

      // Now clear selection and apply again
      ordersWin.selectedCells = new Set();
      ordersWin.anchorCell = null;
      app._test.clearLinkFilters(ordersWin);
      const afterClear = Object.keys(custWin.linkFilters).length;

      return { hadFilter, afterClear };
    });

    expect(result.hadFilter).toBe(true);
    expect(result.afterClear).toBe(0);
  });

  test('link with regex column matches multiple columns', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Regex Col',
      links: [
        { source: { table: 'customers', column: '(id|name)' }, target: { table: 'orders', column: 'customer_id' } }
      ]
    });

    const result = await page.evaluate(() => {
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custTable = app._test.tables['customers'];
      const row = custTable.rows[0]; // id=1, name=Alice

      custWin.selectedCells = new Set();
      for (const col of custTable.columns) {
        custWin.selectedCells.add(`${row._rownum}:${col}`);
      }
      custWin.anchorCell = { rownum: row._rownum, col: 'id' };
      app._test.applyLinkFilters(custWin);

      return ordersWin.linkFilters['customer_id']
        ? [...ordersWin.linkFilters['customer_id']].sort()
        : [];
    });

    // Source columns "id" and "name" both match — values are "1" and "Alice"
    expect(result).toEqual(['1', 'Alice']);
  });
});
