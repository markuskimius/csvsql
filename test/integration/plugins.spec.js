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

  test('format chip appears in status bar when plugins match', async ({ page }) => {
    let chip = await page.$('.status-chip-format');
    expect(chip).toBeNull();

    await loadPluginConfig(page, {
      name: 'Toggle Test', table: '.*',
      columns: [{ match: '^name$', display: 'upper(value)' }]
    });
    await page.evaluate(() => app._test.rerenderAllWindows());
    await page.waitForTimeout(100);

    chip = await page.$('.status-chip-format');
    expect(chip).not.toBeNull();
    expect(await chip.evaluate(el => !el.classList.contains('off'))).toBe(true);
  });

  test('format chip does not appear when no plugins match', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'No Match', table: '^nonexistent$',
      columns: [{ match: '^name$', display: 'upper(value)' }]
    });
    await page.evaluate(() => app._test.rerenderAllWindows());
    await page.waitForTimeout(100);

    const chip = await page.$('.status-chip-format');
    expect(chip).toBeNull();
  });

  test('format chip disables all transforms, headers show feature-disabled', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Toggle Disable', table: '.*',
      columns: [{ match: '^name$', display: 'upper(value)' }]
    });
    await page.evaluate(() => app._test.rerenderAllWindows());
    await page.waitForTimeout(100);

    const before = await getCellText(page, 0, 0);
    expect(before).toBe('ALICE JOHNSON');

    let th = await page.$('th.col-transformed');
    expect(th).not.toBeNull();
    expect(await th.evaluate(el => el.classList.contains('feature-disabled'))).toBe(false);

    await page.click('.status-chip-format');
    await page.waitForTimeout(100);

    const after = await getCellText(page, 0, 0);
    expect(after).toBe('Alice Johnson');

    const chip = await page.$('.status-chip-format');
    expect(await chip.evaluate(el => el.classList.contains('off'))).toBe(true);

    th = await page.$('th.col-transformed');
    expect(th).not.toBeNull();
    expect(await th.evaluate(el => el.classList.contains('feature-disabled'))).toBe(true);
  });

  test('format chip re-enables all transforms', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Toggle Re-enable', table: '.*',
      columns: [{ match: '^name$', display: 'upper(value)' }]
    });
    await page.evaluate(() => app._test.rerenderAllWindows());
    await page.waitForTimeout(100);

    await page.click('.status-chip-format');
    await page.waitForTimeout(100);
    await page.click('.status-chip-format');
    await page.waitForTimeout(100);

    const cell = await getCellText(page, 0, 0);
    expect(cell).toBe('ALICE JOHNSON');

    const chip = await page.$('.status-chip-format');
    expect(await chip.evaluate(el => !el.classList.contains('off'))).toBe(true);

    const th = await page.$('th.col-transformed');
    expect(await th.evaluate(el => !el.classList.contains('feature-disabled'))).toBe(true);
  });

  test('disabling a single column transform via disabledTransforms', async ({ page }) => {
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

    const transformedThs = await page.$$('th.col-transformed');
    expect(transformedThs.length).toBe(2);

    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.disabledTransforms.add('name');
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const nameAfter = await getCellText(page, 0, 0);
    expect(nameAfter).toBe('Alice Johnson');

    const emailIdx = await page.evaluate(() => app._test.windows[0]._columns.indexOf('email'));
    const emailCell = await getCellText(page, 0, emailIdx);
    expect(emailCell).toBe('ALICE@EXAMPLE.COM');
  });

  test('re-enabling a disabled column transform via disabledTransforms', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Re-enable Col', table: '.*',
      columns: [{ match: '^name$', display: 'upper(value)' }]
    });
    await page.evaluate(() => app._test.rerenderAllWindows());
    await page.waitForTimeout(100);

    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.disabledTransforms.add('name');
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);
    expect(await getCellText(page, 0, 0)).toBe('Alice Johnson');

    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.disabledTransforms.delete('name');
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);
    expect(await getCellText(page, 0, 0)).toBe('ALICE JOHNSON');

    const th = await page.$('th.col-transformed');
    expect(await th.evaluate(el => !el.classList.contains('feature-disabled'))).toBe(true);
  });

  test('col-transformed class only appears on transformed columns', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Selective', table: '.*',
      columns: [{ match: '^name$', display: 'upper(value)' }]
    });
    await page.evaluate(() => app._test.rerenderAllWindows());
    await page.waitForTimeout(100);

    const transformed = await page.$$eval('th.col-transformed', els => els.map(el => {
      const colName = el.querySelector('.col-name');
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

    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.disabledTransforms.add('name');
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const cell = await getCellText(page, 0, 0);
    expect(cell).toBe('Alice Johnson');

    const th = await page.$('th.col-transformed');
    expect(await th.evaluate(el => el.classList.contains('feature-disabled'))).toBe(true);
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

    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.disabledTransforms.add('name');
      app._test.rebuildTable(win);
    });
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

    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.disabledTransforms.add('name');
      app._test.rebuildTable(win);
    });
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

  test('format chip re-enables all when some columns were disabled', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Multi', table: '.*',
      columns: [
        { match: '^name$', display: 'upper(value)' },
        { match: '^email$', display: 'upper(value)' }
      ]
    });
    await page.evaluate(() => app._test.rerenderAllWindows());
    await page.waitForTimeout(100);

    let chip = await page.$('.status-chip-format');
    expect(await chip.evaluate(el => !el.classList.contains('off'))).toBe(true);

    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.disabledTransforms.add('name');
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    await page.click('.status-chip-format');
    await page.waitForTimeout(100);

    chip = await page.$('.status-chip-format');
    expect(await chip.evaluate(el => el.classList.contains('off'))).toBe(true);
    expect(await getCellText(page, 0, 0)).toBe('Alice Johnson');

    await page.click('.status-chip-format');
    await page.waitForTimeout(100);

    chip = await page.$('.status-chip-format');
    expect(await chip.evaluate(el => !el.classList.contains('off'))).toBe(true);
    expect(await getCellText(page, 0, 0)).toBe('ALICE JOHNSON');
  });

  test('unloading plugin removes transform class and format chip', async ({ page }) => {
    await loadPluginConfig(page, {
      name: 'Removable', table: '.*',
      columns: [{ match: '^name$', display: 'upper(value)' }]
    });
    await page.evaluate(() => app._test.rerenderAllWindows());
    await page.waitForTimeout(100);

    expect(await page.$('th.col-transformed')).not.toBeNull();
    expect(await page.$('.status-chip-format')).not.toBeNull();

    await page.evaluate(() => app._test.unloadPlugin(0));
    await page.waitForTimeout(100);

    expect(await page.$('th.col-transformed')).toBeNull();
    expect(await page.$('.status-chip-format')).toBeNull();
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
    const linkChip = custWindow.locator('.status-chip-link');
    await expect(linkChip).toBeVisible();
    expect(await linkChip.textContent()).toBe('Link');
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

test.describe('Sort chip', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, 'sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  async function getCellText(page, rowIdx, colIdx) {
    return page.evaluate(({ rowIdx, colIdx }) => {
      const rows = document.querySelectorAll('.subwindow table tbody tr:not(.virtual-pad)');
      if (!rows[rowIdx]) return null;
      const tds = rows[rowIdx].querySelectorAll('td.data-cell');
      return tds[colIdx] ? tds[colIdx].textContent : null;
    }, { rowIdx, colIdx });
  }

  test('sort chip appears when column is sorted', async ({ page }) => {
    let chip = await page.$('.status-chip-sort');
    expect(chip).toBeNull();

    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    chip = await page.$('.status-chip-sort');
    expect(chip).not.toBeNull();
    expect(await chip.textContent()).toBe('Sort');
  });

  test('sort chip disappears when sort is cleared', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);
    expect(await page.$('.status-chip-sort')).not.toBeNull();

    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);
    expect(await page.$('.status-chip-sort')).toBeNull();
  });

  test('clicking sort chip suspends sorting and shows .off', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    // Capture sorted order
    const sortedFirst = await getCellText(page, 0, 0);

    // Click chip to suspend
    await page.click('.status-chip-sort');
    await page.waitForTimeout(100);

    const chip = await page.$('.status-chip-sort');
    expect(await chip.evaluate(el => el.classList.contains('off'))).toBe(true);

    // Rows should revert to original order (unsorted)
    const unsortedFirst = await getCellText(page, 0, 0);
    // Original first row is Alice Johnson; sorted asc would also start with Alice
    // Use descending to make the difference clear
    const state = await page.evaluate(() => {
      const win = app._test.windows[0];
      return { disableSort: win.disableSort, sortCols: win.sortCols };
    });
    expect(state.disableSort).toBe(true);
    expect(state.sortCols.length).toBe(1); // sort config preserved
  });

  test('sort chip suspends sorting - rows revert to original order', async ({ page }) => {
    // Sort descending so first row changes from original
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'desc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const sortedFirst = await getCellText(page, 0, 0);
    // Descending by name: last alphabetically should be first
    expect(sortedFirst).not.toBe('Alice Johnson');

    // Suspend sorting
    await page.click('.status-chip-sort');
    await page.waitForTimeout(100);

    const unsortedFirst = await getCellText(page, 0, 0);
    expect(unsortedFirst).toBe('Alice Johnson'); // original order restored
  });

  test('column header keeps .sorted but gains .feature-disabled when sort suspended', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    let th = await page.$('th.sorted');
    expect(th).not.toBeNull();
    expect(await th.evaluate(el => el.classList.contains('feature-disabled'))).toBe(false);

    await page.click('.status-chip-sort');
    await page.waitForTimeout(100);

    th = await page.$('th.sorted');
    expect(th).not.toBeNull();
    expect(await th.evaluate(el => el.classList.contains('feature-disabled'))).toBe(true);
  });

  test('clicking sort chip again re-enables sorting', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'desc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const sortedFirst = await getCellText(page, 0, 0);

    // Suspend
    await page.click('.status-chip-sort');
    await page.waitForTimeout(100);

    // Re-enable
    await page.click('.status-chip-sort');
    await page.waitForTimeout(100);

    const chip = await page.$('.status-chip-sort');
    expect(await chip.evaluate(el => !el.classList.contains('off'))).toBe(true);

    const reEnabledFirst = await getCellText(page, 0, 0);
    expect(reEnabledFirst).toBe(sortedFirst); // same as original sorted order

    const th = await page.$('th.sorted');
    expect(await th.evaluate(el => !el.classList.contains('feature-disabled'))).toBe(true);
  });

  test('sort configuration survives toggle (sort arrows still show)', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    // Suspend and re-enable
    await page.click('.status-chip-sort');
    await page.waitForTimeout(100);
    await page.click('.status-chip-sort');
    await page.waitForTimeout(100);

    // Sort arrows should still be present
    const arrow = await page.$('th.sorted .sort-arrow');
    expect(arrow).not.toBeNull();
    const arrowText = await arrow.textContent();
    expect(arrowText).toContain('▲'); // up arrow for asc

    // sortCols should still be configured
    const sortCols = await page.evaluate(() => app._test.windows[0].sortCols);
    expect(sortCols).toEqual([{ col: 'name', dir: 'asc' }]);
  });
});

test.describe('Filter chip', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, 'sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  async function getVisibleRowCount(page) {
    return page.evaluate(() => {
      return document.querySelectorAll('.subwindow table tbody tr:not(.virtual-pad)').length;
    });
  }

  test('filter chip appears when column autofilter is active', async ({ page }) => {
    let chip = await page.$('.status-chip-filter');
    expect(chip).toBeNull();

    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.columnFilters = { name: new Set(['Alice Johnson']) };
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    chip = await page.$('.status-chip-filter');
    expect(chip).not.toBeNull();
    expect(await chip.textContent()).toBe('Filter');
    expect(await chip.evaluate(el => !el.classList.contains('off'))).toBe(true);
  });

  test('filter chip appears when WHERE filter text is set', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.filterText = "name = 'Alice Johnson'";
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const chip = await page.$('.status-chip-filter');
    expect(chip).not.toBeNull();
    expect(await chip.textContent()).toBe('Filter');
  });

  test('clicking filter chip shows all rows (filter suspended)', async ({ page }) => {
    const totalRows = await getVisibleRowCount(page);

    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.columnFilters = { name: new Set(['Alice Johnson']) };
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const filteredRows = await getVisibleRowCount(page);
    expect(filteredRows).toBe(1);

    // Suspend filter
    await page.click('.status-chip-filter');
    await page.waitForTimeout(100);

    const suspendedRows = await getVisibleRowCount(page);
    expect(suspendedRows).toBe(totalRows);
  });

  test('filter chip shows .off class when suspended', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.columnFilters = { name: new Set(['Alice Johnson']) };
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    await page.click('.status-chip-filter');
    await page.waitForTimeout(100);

    const chip = await page.$('.status-chip-filter');
    expect(await chip.evaluate(el => el.classList.contains('off'))).toBe(true);

    // Column header should have feature-disabled
    const th = await page.$('th.col-filtered');
    expect(th).not.toBeNull();
    expect(await th.evaluate(el => el.classList.contains('feature-disabled'))).toBe(true);
  });

  test('Clear Filters link hidden when filter is suspended', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.columnFilters = { name: new Set(['Alice Johnson']) };
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    // Clear Filters link should be visible
    let clearLink = await page.$('.status-clear-filters');
    expect(clearLink).not.toBeNull();

    // Suspend filter
    await page.click('.status-chip-filter');
    await page.waitForTimeout(100);

    // Clear Filters link should be hidden
    clearLink = await page.$('.status-clear-filters');
    expect(clearLink).toBeNull();
  });

  test('re-enabling filter re-applies the filter', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.columnFilters = { name: new Set(['Alice Johnson']) };
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const filteredRows = await getVisibleRowCount(page);
    expect(filteredRows).toBe(1);

    // Suspend
    await page.click('.status-chip-filter');
    await page.waitForTimeout(100);
    const totalRows = await getVisibleRowCount(page);
    expect(totalRows).toBeGreaterThan(1);

    // Re-enable
    await page.click('.status-chip-filter');
    await page.waitForTimeout(100);

    const reFilteredRows = await getVisibleRowCount(page);
    expect(reFilteredRows).toBe(1);

    const chip = await page.$('.status-chip-filter');
    expect(await chip.evaluate(el => !el.classList.contains('off'))).toBe(true);

    const th = await page.$('th.col-filtered');
    expect(await th.evaluate(el => !el.classList.contains('feature-disabled'))).toBe(true);
  });
});

test.describe('Link chip', () => {
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

  async function selectOrderRow(page, rowIndex) {
    await page.evaluate((idx) => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const ordersTable = app._test.tables['orders'];
      const row = ordersTable.rows[idx];
      ordersWin.selectedCells = new Set();
      for (const col of ordersTable.columns) {
        ordersWin.selectedCells.add(`${row._rownum}:${col}`);
      }
      ordersWin.anchorCell = { rownum: row._rownum, col: 'order_id' };
      app._test.applyLinkFilters(ordersWin);
    }, rowIndex);
  }

  test('link chip appears when link filters are active', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Link Chip Test',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    // Before selection: no link chip on customers
    const custWindow = page.locator('.subwindow').filter({ hasText: 'customers' });
    let linkChip = await custWindow.locator('.status-chip-link').count();
    expect(linkChip).toBe(0);

    await selectOrderRow(page, 0);
    await page.waitForTimeout(200);

    linkChip = await custWindow.locator('.status-chip-link').count();
    expect(linkChip).toBe(1);
    const chipText = await custWindow.locator('.status-chip-link').textContent();
    expect(chipText).toBe('Link');
  });

  test('clicking link chip shows all rows in target (link suspended)', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Link Suspend Test',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    await selectOrderRow(page, 0); // customer_id=1 -> only Alice in customers
    await page.waitForTimeout(200);

    // Verify filter is active: only 1 customer row visible
    const filteredRows = await page.evaluate(() => {
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      return custWin._displayRows ? custWin._displayRows.length : -1;
    });
    expect(filteredRows).toBe(1);

    // Click link chip to suspend (use evaluate to avoid occlusion issues)
    await page.evaluate(() => {
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      const chip = custWin.statusbarEl.querySelector('.status-chip-link');
      chip.click();
    });
    await page.waitForTimeout(200);

    const suspendedRows = await page.evaluate(() => {
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      return custWin._displayRows ? custWin._displayRows.length : -1;
    });
    expect(suspendedRows).toBe(3); // all 3 customers visible
  });

  test('link chip shows .off class when suspended', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Link Off Test',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    await selectOrderRow(page, 0);
    await page.waitForTimeout(200);

    // Click link chip to suspend via evaluate
    const result = await page.evaluate(() => {
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      const chip = custWin.statusbarEl.querySelector('.status-chip-link');
      chip.click();
      // Re-query after rebuild
      const chipAfter = custWin.statusbarEl.querySelector('.status-chip-link');
      const th = custWin.el.querySelector('th.col-linked');
      return {
        chipOff: chipAfter ? chipAfter.classList.contains('off') : null,
        thDisabled: th ? th.classList.contains('feature-disabled') : null,
      };
    });

    expect(result.chipOff).toBe(true);
    expect(result.thDisabled).toBe(true);
  });

  test('re-enabling link chip re-applies link filter', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Link Re-enable Test',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    await selectOrderRow(page, 0);
    await page.waitForTimeout(200);

    // Suspend
    await page.evaluate(() => {
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      custWin.statusbarEl.querySelector('.status-chip-link').click();
    });
    await page.waitForTimeout(100);
    const suspendedRows = await page.evaluate(() => {
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      return custWin._displayRows.length;
    });
    expect(suspendedRows).toBe(3);

    // Re-enable
    await page.evaluate(() => {
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      custWin.statusbarEl.querySelector('.status-chip-link').click();
    });
    await page.waitForTimeout(100);

    const result = await page.evaluate(() => {
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      const chip = custWin.statusbarEl.querySelector('.status-chip-link');
      return {
        rows: custWin._displayRows.length,
        chipOff: chip ? chip.classList.contains('off') : null,
      };
    });
    expect(result.rows).toBe(1);
    expect(result.chipOff).toBe(false);
  });
});

test.describe('Status chip keyboard shortcuts', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, 'sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  async function getVisibleRowCount(page) {
    return page.evaluate(() => {
      return document.querySelectorAll('.subwindow table tbody tr:not(.virtual-pad)').length;
    });
  }

  test('Ctrl+Shift+1 toggles sort', async ({ page }) => {
    // Set up sort
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'desc' }];
      app._test.rebuildTable(win);
      app._test.focusWindow(win.id);
    });
    await page.waitForTimeout(100);

    // Focus a cell so the window is active
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.waitForTimeout(100);

    const modifier = process.platform === 'darwin' ? 'Meta' : 'Control';
    await page.keyboard.press(`${modifier}+Shift+1`);
    await page.waitForTimeout(100);

    // Sort should be disabled
    const state = await page.evaluate(() => {
      const win = app._test.windows[0];
      return { disableSort: win.disableSort };
    });
    expect(state.disableSort).toBe(true);

    const chip = await page.$('.status-chip-sort');
    expect(chip).not.toBeNull();
    expect(await chip.evaluate(el => el.classList.contains('off'))).toBe(true);

    // Toggle back
    await page.keyboard.press(`${modifier}+Shift+1`);
    await page.waitForTimeout(100);

    const stateAfter = await page.evaluate(() => {
      const win = app._test.windows[0];
      return { disableSort: win.disableSort };
    });
    expect(stateAfter.disableSort).toBe(false);
  });

  test('Ctrl+Shift+2 toggles filter', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.columnFilters = { name: new Set(['Alice Johnson']) };
      app._test.rebuildTable(win);
      app._test.focusWindow(win.id);
    });
    await page.waitForTimeout(100);

    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.waitForTimeout(100);

    const filteredRows = await getVisibleRowCount(page);
    expect(filteredRows).toBe(1);

    const modifier = process.platform === 'darwin' ? 'Meta' : 'Control';
    await page.keyboard.press(`${modifier}+Shift+2`);
    await page.waitForTimeout(100);

    const suspendedRows = await getVisibleRowCount(page);
    expect(suspendedRows).toBeGreaterThan(1);

    const chip = await page.$('.status-chip-filter');
    expect(await chip.evaluate(el => el.classList.contains('off'))).toBe(true);
  });

  test('Ctrl+Shift+3 toggles link disable flag', async ({ page }) => {
    // Even without active link filters, the shortcut should toggle the flag
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.waitForTimeout(100);

    const modifier = process.platform === 'darwin' ? 'Meta' : 'Control';
    await page.keyboard.press(`${modifier}+Shift+3`);
    await page.waitForTimeout(100);

    const state = await page.evaluate(() => {
      const win = app._test.windows[0];
      return { disableLink: win.disableLink };
    });
    expect(state.disableLink).toBe(true);

    await page.keyboard.press(`${modifier}+Shift+3`);
    await page.waitForTimeout(100);

    const stateAfter = await page.evaluate(() => {
      const win = app._test.windows[0];
      return { disableLink: win.disableLink };
    });
    expect(stateAfter.disableLink).toBe(false);
  });
});

test.describe('Status chip integration', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, 'sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  async function getCellText(page, rowIdx, colIdx) {
    return page.evaluate(({ rowIdx, colIdx }) => {
      const rows = document.querySelectorAll('.subwindow table tbody tr:not(.virtual-pad)');
      if (!rows[rowIdx]) return null;
      const tds = rows[rowIdx].querySelectorAll('td.data-cell');
      return tds[colIdx] ? tds[colIdx].textContent : null;
    }, { rowIdx, colIdx });
  }

  async function getVisibleRowCount(page) {
    return page.evaluate(() => {
      return document.querySelectorAll('.subwindow table tbody tr:not(.virtual-pad)').length;
    });
  }

  test('multiple chips visible simultaneously (sort + filter)', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }];
      win.columnFilters = { name: new Set(['Alice Johnson', 'Bob Smith']) };
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const sortChip = await page.$('.status-chip-sort');
    const filterChip = await page.$('.status-chip-filter');
    expect(sortChip).not.toBeNull();
    expect(filterChip).not.toBeNull();

    // Both should be active (not off)
    expect(await sortChip.evaluate(el => !el.classList.contains('off'))).toBe(true);
    expect(await filterChip.evaluate(el => !el.classList.contains('off'))).toBe(true);
  });

  test('disabling filter does not affect sort, and vice versa', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'desc' }];
      win.columnFilters = { name: new Set(['Alice Johnson', 'Bob Smith', 'Eve Davis']) };
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const filteredRows = await getVisibleRowCount(page);
    expect(filteredRows).toBe(3);

    // Suspend filter only
    await page.click('.status-chip-filter');
    await page.waitForTimeout(100);

    // Filter chip should be off, sort chip still on
    const filterChip = await page.$('.status-chip-filter');
    const sortChip = await page.$('.status-chip-sort');
    expect(await filterChip.evaluate(el => el.classList.contains('off'))).toBe(true);
    expect(await sortChip.evaluate(el => !el.classList.contains('off'))).toBe(true);

    // All rows should show (filter suspended) but still sorted desc
    const allRows = await getVisibleRowCount(page);
    expect(allRows).toBe(10); // all 10 rows in sample1.csv
    const firstCell = await getCellText(page, 0, 0);
    expect(firstCell).not.toBe('Alice Johnson'); // still sorted desc

    // Suspend sort
    await page.click('.status-chip-sort');
    await page.waitForTimeout(100);

    // Now first row should be original order
    const unsortedFirst = await getCellText(page, 0, 0);
    expect(unsortedFirst).toBe('Alice Johnson');

    // Both chips should be off
    expect(await page.$eval('.status-chip-sort', el => el.classList.contains('off'))).toBe(true);
    expect(await page.$eval('.status-chip-filter', el => el.classList.contains('off'))).toBe(true);
  });

  test('chips persist across table rebuild (e.g., after cell edit)', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }];
      win.disableSort = true;
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    let chip = await page.$('.status-chip-sort');
    expect(chip).not.toBeNull();
    expect(await chip.evaluate(el => el.classList.contains('off'))).toBe(true);

    // Simulate a cell edit that triggers rebuildTable
    await page.evaluate(() => {
      const win = app._test.windows[0];
      const t = app._test.tables[win.tableName];
      t.rows[0].name = 'Edited Name';
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    // Sort chip should still be present and still off
    chip = await page.$('.status-chip-sort');
    expect(chip).not.toBeNull();
    expect(await chip.evaluate(el => el.classList.contains('off'))).toBe(true);

    // disableSort should still be true
    const state = await page.evaluate(() => app._test.windows[0].disableSort);
    expect(state).toBe(true);
  });
});
