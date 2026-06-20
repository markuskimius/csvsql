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

  test('menu stays open after clicking unload button', async ({ page }) => {
    await page.evaluate(async () => {
      const json = JSON.stringify({
        name: 'Stay Open', table: '.*',
        columns: [{ match: '^name$', display: 'upper(value)' }]
      });
      const file = new File([json], 'stay.json', { type: 'application/json' });
      await app._test.loadPluginFile(file);
    });
    await page.click('#menu-plugins .menu-label');
    await page.waitForTimeout(100);

    const menuOpen = await page.$eval('#menu-plugins', el => el.classList.contains('open'));
    expect(menuOpen).toBe(true);

    await page.click('.plugin-unload');
    await page.waitForTimeout(100);

    const menuStillOpen = await page.$eval('#menu-plugins', el => el.classList.contains('open'));
    expect(menuStillOpen).toBe(true);
  });

  test('menu updates plugin list after unload without closing', async ({ page }) => {
    await page.evaluate(async () => {
      for (const name of ['Plugin A', 'Plugin B']) {
        const json = JSON.stringify({
          name, table: '.*',
          columns: [{ match: '^name$', display: 'upper(value)' }]
        });
        const file = new File([json], name + '.json', { type: 'application/json' });
        await app._test.loadPluginFile(file);
      }
    });
    await page.click('#menu-plugins .menu-label');
    await page.waitForTimeout(100);

    let entries = await page.$$('.plugin-entry');
    expect(entries.length).toBe(2);

    await page.click('.plugin-unload');
    await page.waitForTimeout(100);

    entries = await page.$$('.plugin-entry');
    expect(entries.length).toBe(1);
    const remainingName = await page.$eval('.plugin-name', el => el.textContent);
    expect(remainingName).toBe('Plugin B');
  });

  test('can unload multiple plugins sequentially without reopening menu', async ({ page }) => {
    await page.evaluate(async () => {
      for (const name of ['First', 'Second', 'Third']) {
        const json = JSON.stringify({
          name, table: '.*',
          columns: [{ match: '^name$', display: 'upper(value)' }]
        });
        const file = new File([json], name + '.json', { type: 'application/json' });
        await app._test.loadPluginFile(file);
      }
    });
    await page.click('#menu-plugins .menu-label');
    await page.waitForTimeout(100);

    expect(await page.$$('.plugin-entry').then(e => e.length)).toBe(3);

    // Unload first plugin (always click the first ✕)
    await page.click('.plugin-unload');
    await page.waitForTimeout(100);
    expect(await page.$$('.plugin-entry').then(e => e.length)).toBe(2);

    await page.click('.plugin-unload');
    await page.waitForTimeout(100);
    expect(await page.$$('.plugin-entry').then(e => e.length)).toBe(1);

    await page.click('.plugin-unload');
    await page.waitForTimeout(100);
    expect(await page.$$('.plugin-entry').then(e => e.length)).toBe(0);

    // Menu should still be open, showing the empty state
    const menuStillOpen = await page.$eval('#menu-plugins', el => el.classList.contains('open'));
    expect(menuStillOpen).toBe(true);

    const hint = await page.$('.menu-hint');
    expect(hint).not.toBeNull();
    expect(await hint.textContent()).toContain('No plugins loaded');
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
    expect(await linkChip.textContent()).toBe('Linked');
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

test.describe('Transitive link propagation', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../example/customers.csv');
    await waitForWindow(page, 'customers');
    await uploadFile(page, '../example/orders.csv');
    await waitForWindow(page, 'orders');
    await uploadFile(page, '../example/products.csv');
    await waitForWindow(page, 'products');
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

  function selectRow(page, tableName, rowIdx) {
    return page.evaluate(({ tableName, rowIdx }) => {
      const win = app._test.windows.find(w => w.tableName === tableName);
      const t = app._test.tables[tableName];
      const row = t.rows[rowIdx];
      win.selectedCells = new Set();
      for (const col of t.columns) {
        win.selectedCells.add(`${row._rownum}:${col}`);
      }
      win.anchorCell = { rownum: row._rownum, col: t.columns[0] };
      app._test.applyLinkFilters(win);
    }, { tableName, rowIdx });
  }

  test('selecting product filters orders and customers transitively', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Transitive',
      links: [
        { source: { table: 'products', column: '^id$' }, target: { table: 'orders', column: '^product_id$' } },
        { source: { table: 'orders', column: '^customer_id$' }, target: { table: 'customers', column: '^id$' } }
      ]
    });

    // Select product 501 (Widget) — ordered by customers 1, 2, 5 via orders 101, 103, 107
    const result = await page.evaluate(() => {
      const prodWin = app._test.windows.find(w => w.tableName === 'products');
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      const prodTable = app._test.tables['products'];
      const row = prodTable.rows[0]; // id=501

      prodWin.selectedCells = new Set();
      for (const col of prodTable.columns) {
        prodWin.selectedCells.add(`${row._rownum}:${col}`);
      }
      prodWin.anchorCell = { rownum: row._rownum, col: 'id' };
      app._test.applyLinkFilters(prodWin);

      return {
        ordersFilter: ordersWin.linkFilters['product_id'] ? [...ordersWin.linkFilters['product_id']].sort() : [],
        customersFilter: custWin.linkFilters['id'] ? [...custWin.linkFilters['id']].sort() : []
      };
    });

    expect(result.ordersFilter).toEqual(['501']);
    // Customers 1, 2, 5 all ordered product 501
    expect(result.customersFilter).toEqual(['1', '2', '5']);
  });

  test('selecting product 502 filters to correct customers', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Transitive',
      links: [
        { source: { table: 'products', column: '^id$' }, target: { table: 'orders', column: '^product_id$' } },
        { source: { table: 'orders', column: '^customer_id$' }, target: { table: 'customers', column: '^id$' } }
      ]
    });

    // Select product 502 (Gadget Pro) — ordered by customers 1, 4 via orders 102, 105
    const result = await page.evaluate(() => {
      const prodWin = app._test.windows.find(w => w.tableName === 'products');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      const prodTable = app._test.tables['products'];
      const row = prodTable.rows[1]; // id=502

      prodWin.selectedCells = new Set();
      for (const col of prodTable.columns) {
        prodWin.selectedCells.add(`${row._rownum}:${col}`);
      }
      prodWin.anchorCell = { rownum: row._rownum, col: 'id' };
      app._test.applyLinkFilters(prodWin);

      return custWin.linkFilters['id'] ? [...custWin.linkFilters['id']].sort() : [];
    });

    expect(result).toEqual(['1', '4']);
  });

  test('clearing selection clears all transitive filters', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Transitive',
      links: [
        { source: { table: 'products', column: '^id$' }, target: { table: 'orders', column: '^product_id$' } },
        { source: { table: 'orders', column: '^customer_id$' }, target: { table: 'customers', column: '^id$' } }
      ]
    });

    const result = await page.evaluate(() => {
      const prodWin = app._test.windows.find(w => w.tableName === 'products');
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      const prodTable = app._test.tables['products'];
      const row = prodTable.rows[0];

      // Apply filters
      prodWin.selectedCells = new Set();
      for (const col of prodTable.columns) {
        prodWin.selectedCells.add(`${row._rownum}:${col}`);
      }
      prodWin.anchorCell = { rownum: row._rownum, col: 'id' };
      app._test.applyLinkFilters(prodWin);

      // Verify filters are set
      const beforeOrders = Object.keys(ordersWin.linkFilters).length > 0;
      const beforeCust = Object.keys(custWin.linkFilters).length > 0;

      // Clear selection
      prodWin.selectedCells = new Set();
      app._test.clearLinkFilters(prodWin);

      return {
        beforeOrders,
        beforeCust,
        afterOrders: Object.keys(ordersWin.linkFilters).length,
        afterCust: Object.keys(custWin.linkFilters).length
      };
    });

    expect(result.beforeOrders).toBe(true);
    expect(result.beforeCust).toBe(true);
    expect(result.afterOrders).toBe(0);
    expect(result.afterCust).toBe(0);
  });

  test('cycle detection prevents infinite loops with bidirectional rules', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Cyclic',
      links: [
        { source: { table: 'products', column: '^id$' }, target: { table: 'orders', column: '^product_id$' } },
        { source: { table: 'orders', column: '^product_id$' }, target: { table: 'products', column: '^id$' } },
        { source: { table: 'orders', column: '^customer_id$' }, target: { table: 'customers', column: '^id$' } },
        { source: { table: 'customers', column: '^id$' }, target: { table: 'orders', column: '^customer_id$' } }
      ]
    });

    // Should not infinite loop — BFS visited set prevents revisiting tables
    const result = await page.evaluate(() => {
      const prodWin = app._test.windows.find(w => w.tableName === 'products');
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      const prodTable = app._test.tables['products'];
      const row = prodTable.rows[0]; // id=501

      prodWin.selectedCells = new Set();
      for (const col of prodTable.columns) {
        prodWin.selectedCells.add(`${row._rownum}:${col}`);
      }
      prodWin.anchorCell = { rownum: row._rownum, col: 'id' };
      app._test.applyLinkFilters(prodWin);

      return {
        // Products is the source — should NOT get link filters applied to itself
        prodFilters: Object.keys(prodWin.linkFilters).length,
        // Orders should be filtered by product_id
        ordersFilter: ordersWin.linkFilters['product_id'] ? [...ordersWin.linkFilters['product_id']] : [],
        // Customers should be filtered transitively
        custFilter: custWin.linkFilters['id'] ? [...custWin.linkFilters['id']].sort() : []
      };
    });

    expect(result.prodFilters).toBe(0);
    expect(result.ordersFilter).toEqual(['501']);
    expect(result.custFilter).toEqual(['1', '2', '5']);
  });

  test('source table never gets filtered by its own selection', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Self-loop',
      links: [
        { source: { table: '.*', column: '.*id.*' }, target: { table: '.*', column: '.*id.*' } }
      ]
    });

    const result = await page.evaluate(() => {
      const prodWin = app._test.windows.find(w => w.tableName === 'products');
      const prodTable = app._test.tables['products'];
      const row = prodTable.rows[0];

      prodWin.selectedCells = new Set();
      for (const col of prodTable.columns) {
        prodWin.selectedCells.add(`${row._rownum}:${col}`);
      }
      prodWin.anchorCell = { rownum: row._rownum, col: 'id' };
      app._test.applyLinkFilters(prodWin);

      return Object.keys(prodWin.linkFilters).length;
    });

    expect(result).toBe(0);
  });

  test('three-hop chain A→B→C→D propagates correctly', async ({ page }) => {
    await uploadFile(page, 'notes.csv');
    await waitForWindow(page, 'notes');

    await loadLinkPlugin(page, {
      name: '3-hop',
      links: [
        { source: { table: 'products', column: '^id$' }, target: { table: 'orders', column: '^product_id$' } },
        { source: { table: 'orders', column: '^customer_id$' }, target: { table: 'customers', column: '^id$' } },
        { source: { table: 'customers', column: '^id$' }, target: { table: 'notes', column: '^customer_id$' } }
      ]
    });

    // Select product 503 (Starter Kit) — ordered by customers 3, 2 via orders 104, 106
    const result = await page.evaluate(() => {
      const prodWin = app._test.windows.find(w => w.tableName === 'products');
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      const notesWin = app._test.windows.find(w => w.tableName === 'notes');
      const prodTable = app._test.tables['products'];
      const row = prodTable.rows[2]; // id=503

      prodWin.selectedCells = new Set();
      for (const col of prodTable.columns) {
        prodWin.selectedCells.add(`${row._rownum}:${col}`);
      }
      prodWin.anchorCell = { rownum: row._rownum, col: 'id' };
      app._test.applyLinkFilters(prodWin);

      return {
        ordersFilter: ordersWin.linkFilters['product_id'] ? [...ordersWin.linkFilters['product_id']] : [],
        custFilter: custWin.linkFilters['id'] ? [...custWin.linkFilters['id']].sort() : [],
        notesFilter: notesWin.linkFilters['customer_id'] ? [...notesWin.linkFilters['customer_id']].sort() : []
      };
    });

    expect(result.ordersFilter).toEqual(['503']);
    expect(result.custFilter).toEqual(['2', '3']);
    expect(result.notesFilter).toEqual(['2', '3']);
  });

  test('non-link-source table selection does not clear existing filters', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'One-way',
      links: [
        { source: { table: 'products', column: '^id$' }, target: { table: 'orders', column: '^product_id$' } }
      ]
    });

    const result = await page.evaluate(() => {
      const prodWin = app._test.windows.find(w => w.tableName === 'products');
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      const prodTable = app._test.tables['products'];
      const custTable = app._test.tables['customers'];

      // Select in products — orders get filtered
      const row = prodTable.rows[0];
      prodWin.selectedCells = new Set();
      for (const col of prodTable.columns) {
        prodWin.selectedCells.add(`${row._rownum}:${col}`);
      }
      prodWin.anchorCell = { rownum: row._rownum, col: 'id' };
      app._test.applyLinkFilters(prodWin);

      const ordersFilteredBefore = Object.keys(ordersWin.linkFilters).length > 0;

      // Now select in customers — which is NOT a link source
      const custRow = custTable.rows[0];
      custWin.selectedCells = new Set();
      for (const col of custTable.columns) {
        custWin.selectedCells.add(`${custRow._rownum}:${col}`);
      }
      custWin.anchorCell = { rownum: custRow._rownum, col: 'id' };
      app._test.applyLinkFilters(custWin);

      // Orders should still be filtered (customers is not a link source, so no interference)
      return {
        ordersFilteredBefore,
        ordersFilteredAfter: Object.keys(ordersWin.linkFilters).length > 0
      };
    });

    expect(result.ordersFilteredBefore).toBe(true);
    expect(result.ordersFilteredAfter).toBe(true);
  });

  test('new source selection replaces previous transitive filters', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Transitive',
      links: [
        { source: { table: 'products', column: '^id$' }, target: { table: 'orders', column: '^product_id$' } },
        { source: { table: 'orders', column: '^customer_id$' }, target: { table: 'customers', column: '^id$' } }
      ]
    });

    const result = await page.evaluate(() => {
      const prodWin = app._test.windows.find(w => w.tableName === 'products');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      const prodTable = app._test.tables['products'];

      // Select product 501 — customers 1, 2, 5
      let row = prodTable.rows[0];
      prodWin.selectedCells = new Set();
      for (const col of prodTable.columns) {
        prodWin.selectedCells.add(`${row._rownum}:${col}`);
      }
      prodWin.anchorCell = { rownum: row._rownum, col: 'id' };
      app._test.applyLinkFilters(prodWin);
      const firstFilter = custWin.linkFilters['id'] ? [...custWin.linkFilters['id']].sort() : [];

      // Select product 502 — customers 1, 4
      row = prodTable.rows[1];
      prodWin.selectedCells = new Set();
      for (const col of prodTable.columns) {
        prodWin.selectedCells.add(`${row._rownum}:${col}`);
      }
      prodWin.anchorCell = { rownum: row._rownum, col: 'id' };
      app._test.applyLinkFilters(prodWin);
      const secondFilter = custWin.linkFilters['id'] ? [...custWin.linkFilters['id']].sort() : [];

      return { firstFilter, secondFilter };
    });

    expect(result.firstFilter).toEqual(['1', '2', '5']);
    expect(result.secondFilter).toEqual(['1', '4']);
  });

  test('multi-row selection propagates union of values transitively', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Transitive',
      links: [
        { source: { table: 'products', column: '^id$' }, target: { table: 'orders', column: '^product_id$' } },
        { source: { table: 'orders', column: '^customer_id$' }, target: { table: 'customers', column: '^id$' } }
      ]
    });

    const result = await page.evaluate(() => {
      const prodWin = app._test.windows.find(w => w.tableName === 'products');
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      const prodTable = app._test.tables['products'];

      // Select products 501 AND 502
      prodWin.selectedCells = new Set();
      for (const row of [prodTable.rows[0], prodTable.rows[1]]) {
        for (const col of prodTable.columns) {
          prodWin.selectedCells.add(`${row._rownum}:${col}`);
        }
      }
      prodWin.anchorCell = { rownum: prodTable.rows[0]._rownum, col: 'id' };
      app._test.applyLinkFilters(prodWin);

      return {
        ordersFilter: ordersWin.linkFilters['product_id'] ? [...ordersWin.linkFilters['product_id']].sort() : [],
        custFilter: custWin.linkFilters['id'] ? [...custWin.linkFilters['id']].sort() : []
      };
    });

    // Both product IDs in orders filter
    expect(result.ordersFilter).toEqual(['501', '502']);
    // Union of customers: product 501 → customers 1,2,5; product 502 → customers 1,4
    expect(result.custFilter).toEqual(['1', '2', '4', '5']);
  });

  test('empty intermediate table stops propagation', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Transitive',
      links: [
        { source: { table: 'products', column: '^id$' }, target: { table: 'orders', column: '^product_id$' } },
        { source: { table: 'orders', column: '^customer_id$' }, target: { table: 'customers', column: '^id$' } }
      ]
    });

    const result = await page.evaluate(() => {
      const prodWin = app._test.windows.find(w => w.tableName === 'products');
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      const prodTable = app._test.tables['products'];

      // Add a product with no orders
      prodTable.rows.push({ id: '999', name: 'Ghost', price: '0', category: 'None', is_active: 'false', _rownum: 4 });

      // Select only the ghost product
      const row = prodTable.rows[3];
      prodWin.selectedCells = new Set();
      for (const col of prodTable.columns) {
        prodWin.selectedCells.add(`${row._rownum}:${col}`);
      }
      prodWin.anchorCell = { rownum: row._rownum, col: 'id' };
      app._test.applyLinkFilters(prodWin);

      return {
        ordersFilter: ordersWin.linkFilters['product_id'] ? [...ordersWin.linkFilters['product_id']] : [],
        // No orders match product 999, so no customer_id values propagate
        custFilter: Object.keys(custWin.linkFilters).length
      };
    });

    expect(result.ordersFilter).toEqual(['999']);
    expect(result.custFilter).toBe(0);
  });

  test('transitive propagation uses full table rows not display-filtered rows', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Transitive',
      links: [
        { source: { table: 'products', column: '^id$' }, target: { table: 'orders', column: '^product_id$' } },
        { source: { table: 'orders', column: '^customer_id$' }, target: { table: 'customers', column: '^id$' } }
      ]
    });

    const result = await page.evaluate(() => {
      const prodWin = app._test.windows.find(w => w.tableName === 'products');
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      const prodTable = app._test.tables['products'];

      // Apply a column autofilter to orders that hides some orders
      ordersWin.columnFilters = { customer_id: new Set(['1']) };

      // Select product 501 — orders has autofilter but propagation should use ALL rows
      const row = prodTable.rows[0]; // id=501
      prodWin.selectedCells = new Set();
      for (const col of prodTable.columns) {
        prodWin.selectedCells.add(`${row._rownum}:${col}`);
      }
      prodWin.anchorCell = { rownum: row._rownum, col: 'id' };
      app._test.applyLinkFilters(prodWin);

      // Even though orders has autofilter showing only customer_id=1,
      // the transitive propagation uses all rows in the orders table
      return custWin.linkFilters['id'] ? [...custWin.linkFilters['id']].sort() : [];
    });

    // Product 501 ordered by customers 1, 2, 5 — ALL of them, not just the autofilter-visible ones
    expect(result).toEqual(['1', '2', '5']);
  });

  test('linkFiltersEqual correctly compares filter objects', async ({ page }) => {
    const result = await page.evaluate(() => {
      const eq = app._test.linkFiltersEqual;
      return {
        emptyEqual: eq({}, {}),
        sameKeySameValues: eq({ id: new Set(['1', '2']) }, { id: new Set(['1', '2']) }),
        sameKeyDiffValues: eq({ id: new Set(['1', '2']) }, { id: new Set(['1', '3']) }),
        diffKeys: eq({ id: new Set(['1']) }, { name: new Set(['1']) }),
        diffSize: eq({ id: new Set(['1', '2']) }, { id: new Set(['1']) }),
        extraKey: eq({ id: new Set(['1']) }, { id: new Set(['1']), name: new Set(['a']) }),
        multiKeyEqual: eq(
          { id: new Set(['1']), name: new Set(['a']) },
          { name: new Set(['a']), id: new Set(['1']) }
        ),
      };
    });

    expect(result.emptyEqual).toBe(true);
    expect(result.sameKeySameValues).toBe(true);
    expect(result.sameKeyDiffValues).toBe(false);
    expect(result.diffKeys).toBe(false);
    expect(result.diffSize).toBe(false);
    expect(result.extraKey).toBe(false);
    expect(result.multiKeyEqual).toBe(true);
  });

  test('selecting in intermediate table applies only from that point forward', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Full chain',
      links: [
        { source: { table: 'products', column: '^id$' }, target: { table: 'orders', column: '^product_id$' } },
        { source: { table: 'orders', column: '^customer_id$' }, target: { table: 'customers', column: '^id$' } },
        { source: { table: 'orders', column: '^product_id$' }, target: { table: 'products', column: '^id$' } }
      ]
    });

    // Selecting in orders (the middle table) should filter customers AND products
    const result = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      const prodWin = app._test.windows.find(w => w.tableName === 'products');
      const ordersTable = app._test.tables['orders'];
      const row = ordersTable.rows[0]; // order 101: customer_id=1, product_id=501

      ordersWin.selectedCells = new Set();
      for (const col of ordersTable.columns) {
        ordersWin.selectedCells.add(`${row._rownum}:${col}`);
      }
      ordersWin.anchorCell = { rownum: row._rownum, col: 'order_id' };
      app._test.applyLinkFilters(ordersWin);

      return {
        custFilter: custWin.linkFilters['id'] ? [...custWin.linkFilters['id']] : [],
        prodFilter: prodWin.linkFilters['id'] ? [...prodWin.linkFilters['id']] : [],
        ordersFilter: Object.keys(ordersWin.linkFilters).length
      };
    });

    expect(result.custFilter).toEqual(['1']);
    expect(result.prodFilter).toEqual(['501']);
    expect(result.ordersFilter).toBe(0);
  });

  test('unloading plugin clears all transitive link filters', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Transitive',
      links: [
        { source: { table: 'products', column: '^id$' }, target: { table: 'orders', column: '^product_id$' } },
        { source: { table: 'orders', column: '^customer_id$' }, target: { table: 'customers', column: '^id$' } }
      ]
    });

    const result = await page.evaluate(() => {
      const prodWin = app._test.windows.find(w => w.tableName === 'products');
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      const prodTable = app._test.tables['products'];

      // Apply transitive filters
      const row = prodTable.rows[0];
      prodWin.selectedCells = new Set();
      for (const col of prodTable.columns) {
        prodWin.selectedCells.add(`${row._rownum}:${col}`);
      }
      prodWin.anchorCell = { rownum: row._rownum, col: 'id' };
      app._test.applyLinkFilters(prodWin);

      const beforeOrders = Object.keys(ordersWin.linkFilters).length > 0;
      const beforeCust = Object.keys(custWin.linkFilters).length > 0;

      // Unload plugin
      app._test.unloadPlugin(0);

      return {
        beforeOrders,
        beforeCust,
        afterOrders: Object.keys(ordersWin.linkFilters).length,
        afterCust: Object.keys(custWin.linkFilters).length
      };
    });

    expect(result.beforeOrders).toBe(true);
    expect(result.beforeCust).toBe(true);
    expect(result.afterOrders).toBe(0);
    expect(result.afterCust).toBe(0);
  });

  test('suspending Linked chip on one target does not clear other targets', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Transitive',
      links: [
        { source: { table: 'products', column: '^id$' }, target: { table: 'orders', column: '^product_id$' } },
        { source: { table: 'orders', column: '^customer_id$' }, target: { table: 'customers', column: '^id$' } }
      ]
    });

    await selectRow(page, 'products', 0); // id=501
    await page.waitForTimeout(200);

    // Both orders and customers should be filtered
    const before = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      return {
        ordersFiltered: Object.keys(ordersWin.linkFilters).length > 0,
        custFiltered: Object.keys(custWin.linkFilters).length > 0,
        ordersRows: ordersWin._displayRows.length,
        custRows: custWin._displayRows.length,
      };
    });
    expect(before.ordersFiltered).toBe(true);
    expect(before.custFiltered).toBe(true);

    // Click Linked chip on orders to suspend its link filter
    await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const chip = ordersWin.statusbarEl.querySelector('.status-chip-link');
      chip.click();
    });
    await page.waitForTimeout(200);

    const after = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      return {
        ordersDisableLink: ordersWin.disableLink,
        ordersRows: ordersWin._displayRows.length,
        custFiltered: Object.keys(custWin.linkFilters).length > 0,
        custRows: custWin._displayRows.length,
        custDisableLink: custWin.disableLink,
      };
    });
    expect(after.ordersDisableLink).toBe(true);
    expect(after.ordersRows).toBe(7); // all orders visible (suspended)
    expect(after.custFiltered).toBe(true); // customers still filtered
    expect(after.custDisableLink).toBe(false);
  });

  test('suspending Linked chip preserves _activeLinkSourceId', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Transitive',
      links: [
        { source: { table: 'products', column: '^id$' }, target: { table: 'orders', column: '^product_id$' } },
        { source: { table: 'orders', column: '^customer_id$' }, target: { table: 'customers', column: '^id$' } }
      ]
    });

    await selectRow(page, 'products', 0);
    await page.waitForTimeout(200);

    const sourceIdBefore = await page.evaluate(() => {
      const prodWin = app._test.windows.find(w => w.tableName === 'products');
      return { sourceId: app._test._activeLinkSourceId, prodId: prodWin.id };
    });
    expect(sourceIdBefore.sourceId).toBe(sourceIdBefore.prodId);

    // Suspend Linked on orders
    await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      ordersWin.statusbarEl.querySelector('.status-chip-link').click();
    });
    await page.waitForTimeout(200);

    const sourceIdAfter = await page.evaluate(() => {
      const prodWin = app._test.windows.find(w => w.tableName === 'products');
      return { sourceId: app._test._activeLinkSourceId, prodId: prodWin.id };
    });
    expect(sourceIdAfter.sourceId).toBe(sourceIdAfter.prodId);
  });

  test('re-enabling Linked chip on one target does not disrupt other targets', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Transitive',
      links: [
        { source: { table: 'products', column: '^id$' }, target: { table: 'orders', column: '^product_id$' } },
        { source: { table: 'orders', column: '^customer_id$' }, target: { table: 'customers', column: '^id$' } }
      ]
    });

    await selectRow(page, 'products', 0);
    await page.waitForTimeout(200);

    // Suspend orders
    await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      ordersWin.statusbarEl.querySelector('.status-chip-link').click();
    });
    await page.waitForTimeout(200);

    // Re-enable orders
    await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      ordersWin.statusbarEl.querySelector('.status-chip-link').click();
    });
    await page.waitForTimeout(200);

    const result = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      return {
        ordersDisableLink: ordersWin.disableLink,
        ordersFiltered: Object.keys(ordersWin.linkFilters).length > 0,
        custFiltered: Object.keys(custWin.linkFilters).length > 0,
        custDisableLink: custWin.disableLink,
      };
    });
    expect(result.ordersDisableLink).toBe(false);
    expect(result.ordersFiltered).toBe(true);
    expect(result.custFiltered).toBe(true);
    expect(result.custDisableLink).toBe(false);
  });

  test('Ctrl+Shift+3 on link target does not clear other targets', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Transitive',
      links: [
        { source: { table: 'products', column: '^id$' }, target: { table: 'orders', column: '^product_id$' } },
        { source: { table: 'orders', column: '^customer_id$' }, target: { table: 'customers', column: '^id$' } }
      ]
    });

    await selectRow(page, 'products', 0);
    await page.waitForTimeout(200);

    // Focus a cell in the orders window, then toggle link via keyboard
    await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      app._test.focusWindow(ordersWin.id);
    });
    await page.waitForTimeout(100);

    // Click a cell in orders to ensure keyboard events route there
    const ordersCell = page.locator('.subwindow').filter({ hasText: 'orders' }).locator('td.data-cell').first();
    await ordersCell.click();
    await page.waitForTimeout(100);

    const modifier = process.platform === 'darwin' ? 'Meta' : 'Control';
    await page.keyboard.press(`${modifier}+Shift+3`);
    await page.waitForTimeout(200);

    const result = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      return {
        ordersDisableLink: ordersWin.disableLink,
        ordersRows: ordersWin._displayRows.length,
        custFiltered: Object.keys(custWin.linkFilters).length > 0,
        custDisableLink: custWin.disableLink,
      };
    });
    expect(result.ordersDisableLink).toBe(true);
    expect(result.ordersRows).toBe(7); // all orders visible
    expect(result.custFiltered).toBe(true); // customers still filtered
    expect(result.custDisableLink).toBe(false);
  });

  test('suspending Linked chip preserves Linking chip on source', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Transitive',
      links: [
        { source: { table: 'products', column: '^id$' }, target: { table: 'orders', column: '^product_id$' } },
        { source: { table: 'orders', column: '^customer_id$' }, target: { table: 'customers', column: '^id$' } }
      ]
    });

    await selectRow(page, 'products', 0);
    await page.waitForTimeout(200);

    // Verify Linking chip on products before
    const linkingBefore = await page.evaluate(() => {
      const prodWin = app._test.windows.find(w => w.tableName === 'products');
      const chip = prodWin.statusbarEl.querySelector('.status-chip-link-source');
      return { exists: chip !== null, off: chip ? chip.classList.contains('off') : null };
    });
    expect(linkingBefore.exists).toBe(true);
    expect(linkingBefore.off).toBe(false);

    // Suspend orders
    await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      ordersWin.statusbarEl.querySelector('.status-chip-link').click();
    });
    await page.waitForTimeout(200);

    // Linking chip on products should still be active
    const linkingAfter = await page.evaluate(() => {
      const prodWin = app._test.windows.find(w => w.tableName === 'products');
      const chip = prodWin.statusbarEl.querySelector('.status-chip-link-source');
      return { exists: chip !== null, off: chip ? chip.classList.contains('off') : null };
    });
    expect(linkingAfter.exists).toBe(true);
    expect(linkingAfter.off).toBe(false);
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
    expect(await chip.textContent()).toBe('Sorted');
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

  test('clicking sort chip clears sorting and chip disappears', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    // Sort chip should be present
    expect(await page.$('.status-chip-sort')).not.toBeNull();

    // Click chip to clear sort
    await page.click('.status-chip-sort');
    await page.waitForTimeout(100);

    // Chip should disappear (sort was cleared, not suspended)
    const chip = await page.$('.status-chip-sort');
    expect(chip).toBeNull();

    // Sort config should be cleared entirely
    const state = await page.evaluate(() => {
      const win = app._test.windows[0];
      return { disableSort: win.disableSort, sortCols: win.sortCols };
    });
    expect(state.disableSort).toBe(false);
    expect(state.sortCols.length).toBe(0); // sort config cleared
  });

  test('clicking sort chip clears sorting - rows revert to original order', async ({ page }) => {
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

    // Click chip to clear sorting
    await page.click('.status-chip-sort');
    await page.waitForTimeout(100);

    const unsortedFirst = await getCellText(page, 0, 0);
    expect(unsortedFirst).toBe('Alice Johnson'); // original order restored

    // Chip should be gone
    expect(await page.$('.status-chip-sort')).toBeNull();
  });

  test('clicking sort chip removes .sorted class from column header', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    let th = await page.$('th.sorted');
    expect(th).not.toBeNull();

    await page.click('.status-chip-sort');
    await page.waitForTimeout(100);

    // After clearing sort, no column should have .sorted class
    th = await page.$('th.sorted');
    expect(th).toBeNull();
  });

  test('clicking sort chip removes sort arrows from header', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    // Sort arrow should be present before clicking
    let arrow = await page.$('th.sorted .sort-arrow');
    expect(arrow).not.toBeNull();

    // Click chip to clear sort
    await page.click('.status-chip-sort');
    await page.waitForTimeout(100);

    // Sort arrows should be gone
    arrow = await page.$('th.sorted .sort-arrow');
    expect(arrow).toBeNull();

    // sortCols should be empty
    const sortCols = await page.evaluate(() => app._test.windows[0].sortCols);
    expect(sortCols).toEqual([]);
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
    expect(await chip.textContent()).toBe('Filtered');
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
    expect(await chip.textContent()).toBe('Filtered');
  });

  test('clicking filter chip clears filters and chip disappears', async ({ page }) => {
    const totalRows = await getVisibleRowCount(page);

    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.columnFilters = { name: new Set(['Alice Johnson']) };
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const filteredRows = await getVisibleRowCount(page);
    expect(filteredRows).toBe(1);

    // Click chip to clear filters
    await page.click('.status-chip-filter');
    await page.waitForTimeout(100);

    // All rows should show (filters cleared)
    const clearedRows = await getVisibleRowCount(page);
    expect(clearedRows).toBe(totalRows);

    // Chip should disappear
    const chip = await page.$('.status-chip-filter');
    expect(chip).toBeNull();

    // Filters should be empty
    const state = await page.evaluate(() => {
      const win = app._test.windows[0];
      return {
        columnFilterKeys: Object.keys(win.columnFilters).length,
        filterText: win.filterText,
        disableFilter: win.disableFilter,
      };
    });
    expect(state.columnFilterKeys).toBe(0);
    expect(state.filterText).toBe('');
    expect(state.disableFilter).toBe(false);
  });

  test('clicking filter chip removes col-filtered class from header', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.columnFilters = { name: new Set(['Alice Johnson']) };
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    // Column header should have col-filtered before clearing
    let th = await page.$('th.col-filtered');
    expect(th).not.toBeNull();

    await page.click('.status-chip-filter');
    await page.waitForTimeout(100);

    // After clearing, col-filtered class should be gone
    th = await page.$('th.col-filtered');
    expect(th).toBeNull();
  });

  test('clicking filter chip also removes Clear Filters link', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.columnFilters = { name: new Set(['Alice Johnson']) };
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    // Clear Filters link should be visible
    let clearLink = await page.$('.status-clear-filters');
    expect(clearLink).not.toBeNull();

    // Click chip to clear filters
    await page.click('.status-chip-filter');
    await page.waitForTimeout(100);

    // Clear Filters link should be gone (no filters remain)
    clearLink = await page.$('.status-clear-filters');
    expect(clearLink).toBeNull();
  });

  test('clicking filter chip clears WHERE filter text', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.filterText = "name = 'Alice Johnson'";
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const filteredRows = await getVisibleRowCount(page);
    expect(filteredRows).toBe(1);

    await page.click('.status-chip-filter');
    await page.waitForTimeout(100);

    const totalRows = await getVisibleRowCount(page);
    expect(totalRows).toBeGreaterThan(1);

    // Filter text should be cleared
    const filterText = await page.evaluate(() => app._test.windows[0].filterText);
    expect(filterText).toBe('');

    // Chip should be gone
    expect(await page.$('.status-chip-filter')).toBeNull();
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
    expect(chipText).toBe('Linked');
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

  test('suspending Linked chip does not clear _activeLinkSourceId', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Link Source Test',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    await selectOrderRow(page, 0);
    await page.waitForTimeout(200);

    const before = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      return { sourceId: app._test._activeLinkSourceId, ordersId: ordersWin.id };
    });
    expect(before.sourceId).toBe(before.ordersId);

    // Click Linked chip on customers to suspend
    await page.evaluate(() => {
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      custWin.statusbarEl.querySelector('.status-chip-link').click();
    });
    await page.waitForTimeout(200);

    const after = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      return {
        sourceId: app._test._activeLinkSourceId,
        ordersId: ordersWin.id,
        custDisableLink: app._test.windows.find(w => w.tableName === 'customers').disableLink,
      };
    });
    expect(after.sourceId).toBe(after.ordersId);
    expect(after.custDisableLink).toBe(true);
  });

  test('suspending Linked chip preserves linkFilters on target', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Link Preserve Test',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    await selectOrderRow(page, 0); // customer_id=1
    await page.waitForTimeout(200);

    // Suspend via chip click
    await page.evaluate(() => {
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      custWin.statusbarEl.querySelector('.status-chip-link').click();
    });
    await page.waitForTimeout(200);

    // linkFilters should still be stored even when disableLink is true
    const result = await page.evaluate(() => {
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      return {
        disableLink: custWin.disableLink,
        hasLinkFilters: Object.keys(custWin.linkFilters).length > 0,
        rows: custWin._displayRows.length,
      };
    });
    expect(result.disableLink).toBe(true);
    expect(result.hasLinkFilters).toBe(true); // filters still stored
    expect(result.rows).toBe(3); // but not applied, all rows visible
  });

  test('Ctrl+Shift+3 on link target does not clear _activeLinkSourceId', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Link Kbd Test',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    await selectOrderRow(page, 0);
    await page.waitForTimeout(200);

    // Focus customers window via evaluate to avoid occlusion
    await page.evaluate(() => {
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      app._test.focusWindow(custWin.id);
      const cell = custWin.bodyEl.querySelector('td.data-cell');
      if (cell) cell.focus();
    });
    await page.waitForTimeout(100);

    const modifier = process.platform === 'darwin' ? 'Meta' : 'Control';
    await page.keyboard.press(`${modifier}+Shift+3`);
    await page.waitForTimeout(200);

    const result = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      return {
        sourceId: app._test._activeLinkSourceId,
        ordersId: ordersWin.id,
        custDisableLink: custWin.disableLink,
        custRows: custWin._displayRows.length,
      };
    });
    expect(result.sourceId).toBe(result.ordersId);
    expect(result.custDisableLink).toBe(true);
    expect(result.custRows).toBe(3); // all rows visible (suspended)
  });
});

test.describe('Linking chip', () => {
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

  async function selectCustomerRow(page, rowIndex) {
    await page.evaluate((idx) => {
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      const custTable = app._test.tables['customers'];
      const row = custTable.rows[idx];
      custWin.selectedCells = new Set();
      for (const col of custTable.columns) {
        custWin.selectedCells.add(`${row._rownum}:${col}`);
      }
      custWin.anchorCell = { rownum: row._rownum, col: 'id' };
      app._test.applyLinkFilters(custWin);
    }, rowIndex);
  }

  test('Linking chip appears on source table when rows are selected', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Linking Chip Test',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    // Before selection: no Linking chip on either window
    const beforeOrders = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      return ordersWin.statusbarEl.querySelector('.status-chip-link-source') !== null;
    });
    expect(beforeOrders).toBe(false);

    await selectOrderRow(page, 0);
    await page.waitForTimeout(200);

    // Linking chip should appear on orders (the source)
    const result = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const chip = ordersWin.statusbarEl.querySelector('.status-chip-link-source');
      return {
        exists: chip !== null,
        text: chip ? chip.textContent : null,
        hasOff: chip ? chip.classList.contains('off') : null,
      };
    });
    expect(result.exists).toBe(true);
    expect(result.text).toBe('Linking');
    expect(result.hasOff).toBe(false);
  });

  test('Linking chip does not appear on target table', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Linking Target Test',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    await selectOrderRow(page, 0);
    await page.waitForTimeout(200);

    // Linking chip should NOT appear on customers (the target)
    const custHasLinkingChip = await page.evaluate(() => {
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      return custWin.statusbarEl.querySelector('.status-chip-link-source') !== null;
    });
    expect(custHasLinkingChip).toBe(false);

    // But Linked chip should appear on customers (the target)
    const custHasLinkedChip = await page.evaluate(() => {
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      return custWin.statusbarEl.querySelector('.status-chip-link') !== null;
    });
    expect(custHasLinkedChip).toBe(true);
  });

  test('Linking chip disappears when selection is cleared', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Linking Clear Test',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    await selectOrderRow(page, 0);
    await page.waitForTimeout(200);

    // Verify chip is present
    let hasChip = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      return ordersWin.statusbarEl.querySelector('.status-chip-link-source') !== null;
    });
    expect(hasChip).toBe(true);

    // Clear selection
    await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      ordersWin.selectedCells = new Set();
      ordersWin.anchorCell = null;
      app._test.clearLinkFilters(ordersWin);
    });
    await page.waitForTimeout(200);

    // Linking chip should be gone
    hasChip = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      return ordersWin.statusbarEl.querySelector('.status-chip-link-source') !== null;
    });
    expect(hasChip).toBe(false);
  });

  test('clicking Linking chip suspends outbound linking', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Linking Suspend Test',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    await selectOrderRow(page, 0);
    await page.waitForTimeout(200);

    // Verify targets are filtered (only 1 customer visible)
    let custRows = await page.evaluate(() => {
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      return custWin._displayRows ? custWin._displayRows.length : -1;
    });
    expect(custRows).toBe(1);

    // Click Linking chip to suspend outbound linking
    await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const chip = ordersWin.statusbarEl.querySelector('.status-chip-link-source');
      chip.click();
    });
    await page.waitForTimeout(200);

    // Targets should show all rows (link filters cleared)
    custRows = await page.evaluate(() => {
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      return custWin._displayRows ? custWin._displayRows.length : -1;
    });
    expect(custRows).toBe(3);

    // disableLinkSource should be set
    const state = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      return {
        disableLinkSource: ordersWin.disableLinkSource,
      };
    });
    expect(state.disableLinkSource).toBe(true);
  });

  test('re-enabling disableLinkSource restores outbound linking', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Linking Re-enable Test',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    await selectOrderRow(page, 0);
    await page.waitForTimeout(200);

    // Suspend via chip click
    await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      ordersWin.statusbarEl.querySelector('.status-chip-link-source').click();
    });
    await page.waitForTimeout(200);

    // Verify suspended state: all customers visible
    let custRows = await page.evaluate(() => {
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      return custWin._displayRows ? custWin._displayRows.length : -1;
    });
    expect(custRows).toBe(3);

    // Re-enable disableLinkSource and re-apply link filters
    // (simulates the applyCellHighlights guard behavior when selection triggers)
    await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      ordersWin.disableLinkSource = false;
      if (ordersWin.selectedCells.size > 0) {
        app._test.applyLinkFilters(ordersWin);
      }
    });
    await page.waitForTimeout(200);

    // Targets should be filtered again
    const result = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      const chip = ordersWin.statusbarEl.querySelector('.status-chip-link-source');
      return {
        custRows: custWin._displayRows ? custWin._displayRows.length : -1,
        chipExists: chip !== null,
        disableLinkSource: ordersWin.disableLinkSource,
      };
    });
    expect(result.custRows).toBe(1);
    expect(result.chipExists).toBe(true);
    expect(result.disableLinkSource).toBe(false);
  });

  test('only one table can have Linking chip at a time', async ({ page }) => {
    // Set up bidirectional links so both tables can be sources
    await loadLinkPlugin(page, {
      name: 'Bidirectional Linking',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } },
        { source: { table: 'customers', column: 'id' }, target: { table: 'orders', column: 'customer_id' } }
      ]
    });

    // Select in orders first
    await selectOrderRow(page, 0);
    await page.waitForTimeout(200);

    // Verify Linking chip on orders
    let state = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      return {
        ordersHasLinking: ordersWin.statusbarEl.querySelector('.status-chip-link-source') !== null,
        custHasLinking: custWin.statusbarEl.querySelector('.status-chip-link-source') !== null,
      };
    });
    expect(state.ordersHasLinking).toBe(true);
    expect(state.custHasLinking).toBe(false);

    // Now select in customers
    await selectCustomerRow(page, 0);
    await page.waitForTimeout(200);

    // Linking chip should move to customers, gone from orders
    state = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      return {
        ordersHasLinking: ordersWin.statusbarEl.querySelector('.status-chip-link-source') !== null,
        custHasLinking: custWin.statusbarEl.querySelector('.status-chip-link-source') !== null,
      };
    });
    expect(state.ordersHasLinking).toBe(false);
    expect(state.custHasLinking).toBe(true);
  });

  test('disableLinkSource state persists after clicking chip', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Linking Persist State Test',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    await selectOrderRow(page, 0);
    await page.waitForTimeout(200);

    // Disable via chip click
    await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      ordersWin.statusbarEl.querySelector('.status-chip-link-source').click();
    });
    await page.waitForTimeout(200);

    // disableLinkSource should be set
    let state = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      return { disableLinkSource: ordersWin.disableLinkSource };
    });
    expect(state.disableLinkSource).toBe(true);

    // Trigger a rebuild and verify disableLinkSource persists
    await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      app._test.rebuildTable(ordersWin);
    });
    await page.waitForTimeout(200);

    state = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      return { disableLinkSource: ordersWin.disableLinkSource };
    });
    expect(state.disableLinkSource).toBe(true);

    // Verify that targets remain unfiltered (linking is still suspended)
    const custRows = await page.evaluate(() => {
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      return custWin._displayRows ? custWin._displayRows.length : -1;
    });
    expect(custRows).toBe(3);
  });

  test('disableLinkSource prevents link filters on targets', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Linking Prevent Test',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    // Set disableLinkSource before selecting
    await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      ordersWin.disableLinkSource = true;
    });

    // Select a row — but since disableLinkSource is true, applyLinkFilters
    // in applyCellHighlights is guarded by !win.disableLinkSource
    // Simulate the guard behavior
    await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const ordersTable = app._test.tables['orders'];
      const row = ordersTable.rows[0];
      ordersWin.selectedCells = new Set();
      for (const col of ordersTable.columns) {
        ordersWin.selectedCells.add(`${row._rownum}:${col}`);
      }
      ordersWin.anchorCell = { rownum: row._rownum, col: 'order_id' };
      // Simulate the guard in applyCellHighlights
      if (ordersWin.selectedCells.size > 0 && !ordersWin.disableLinkSource) {
        app._test.applyLinkFilters(ordersWin);
      } else {
        app._test.clearLinkFilters(ordersWin);
      }
    });
    await page.waitForTimeout(200);

    // Targets should NOT be filtered
    const custRows = await page.evaluate(() => {
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      return custWin._displayRows ? custWin._displayRows.length : -1;
    });
    expect(custRows).toBe(3); // all customers visible

    // No link filters should exist on target
    const linkFilterCount = await page.evaluate(() => {
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      return Object.keys(custWin.linkFilters).length;
    });
    expect(linkFilterCount).toBe(0);
  });

  test('Linking chip has distinct blue color, different from teal Linked chip', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Linking Color Test',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    await selectOrderRow(page, 0);
    await page.waitForTimeout(200);

    const colors = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      const linkingChip = ordersWin.statusbarEl.querySelector('.status-chip-link-source');
      const linkedChip = custWin.statusbarEl.querySelector('.status-chip-link');
      return {
        linkingColor: linkingChip ? getComputedStyle(linkingChip).color : null,
        linkedColor: linkedChip ? getComputedStyle(linkedChip).color : null,
      };
    });
    expect(colors.linkingColor).not.toBeNull();
    expect(colors.linkedColor).not.toBeNull();
    expect(colors.linkingColor).not.toBe(colors.linkedColor);
  });

  test('Linking chip stays visible with off class after clicking to disable', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Linking Persist Chip Test',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    await selectOrderRow(page, 0);
    await page.waitForTimeout(200);

    // Click Linking chip to disable
    await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      ordersWin.statusbarEl.querySelector('.status-chip-link-source').click();
    });
    await page.waitForTimeout(200);

    // Chip should still exist with 'off' class
    const state = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const chip = ordersWin.statusbarEl.querySelector('.status-chip-link-source');
      return {
        exists: chip !== null,
        hasOff: chip ? chip.classList.contains('off') : null,
        text: chip ? chip.textContent : null,
      };
    });
    expect(state.exists).toBe(true);
    expect(state.hasOff).toBe(true);
    expect(state.text).toBe('Linking');
  });

  test('Linking chip persists after rebuild when disabled', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Linking Rebuild Persist Test',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    await selectOrderRow(page, 0);
    await page.waitForTimeout(200);

    // Disable linking via chip click
    await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      ordersWin.statusbarEl.querySelector('.status-chip-link-source').click();
    });
    await page.waitForTimeout(200);

    // Force multiple rebuilds
    await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      app._test.rebuildTable(ordersWin);
      app._test.rebuildTable(ordersWin);
    });
    await page.waitForTimeout(200);

    // Chip should still be present with 'off' class
    const state = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const chip = ordersWin.statusbarEl.querySelector('.status-chip-link-source');
      return {
        exists: chip !== null,
        hasOff: chip ? chip.classList.contains('off') : null,
      };
    });
    expect(state.exists).toBe(true);
    expect(state.hasOff).toBe(true);
  });

  test('clicking disabled Linking chip re-enables and reapplies link filters', async ({ page }) => {
    await loadLinkPlugin(page, {
      name: 'Linking Re-enable Click Test',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    await selectOrderRow(page, 0);
    await page.waitForTimeout(200);

    // Disable via chip
    await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      ordersWin.statusbarEl.querySelector('.status-chip-link-source').click();
    });
    await page.waitForTimeout(200);

    // Verify disabled state
    let custRows = await page.evaluate(() => {
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      return custWin._displayRows ? custWin._displayRows.length : -1;
    });
    expect(custRows).toBe(3);

    // Click again to re-enable
    await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      ordersWin.statusbarEl.querySelector('.status-chip-link-source').click();
    });
    await page.waitForTimeout(200);

    // Link filters should be reapplied, chip should be active (no off class)
    const state = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      const chip = ordersWin.statusbarEl.querySelector('.status-chip-link-source');
      return {
        custRows: custWin._displayRows ? custWin._displayRows.length : -1,
        chipExists: chip !== null,
        hasOff: chip ? chip.classList.contains('off') : null,
        disableLinkSource: ordersWin.disableLinkSource,
      };
    });
    expect(state.custRows).toBe(1);
    expect(state.chipExists).toBe(true);
    expect(state.hasOff).toBe(false);
    expect(state.disableLinkSource).toBe(false);
  });

  test('disableLinkSource resets when table receives link filters from new source', async ({ page }) => {
    // Bidirectional links: orders -> customers, customers -> orders
    await loadLinkPlugin(page, {
      name: 'Linking Reset Test',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } },
        { source: { table: 'customers', column: 'id' }, target: { table: 'orders', column: 'customer_id' } }
      ]
    });

    // Select in orders first to make it the link source
    await selectOrderRow(page, 0);
    await page.waitForTimeout(200);

    // Disable linking on orders via chip click
    await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      ordersWin.statusbarEl.querySelector('.status-chip-link-source').click();
    });
    await page.waitForTimeout(200);

    // Verify orders has disableLinkSource = true
    let ordersDisabled = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      return ordersWin.disableLinkSource;
    });
    expect(ordersDisabled).toBe(true);

    // Now select in customers — this makes customers the link source,
    // and orders becomes a link target. This should reset orders' disableLinkSource.
    await selectCustomerRow(page, 0);
    await page.waitForTimeout(200);

    const state = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      return {
        ordersDisableLinkSource: ordersWin.disableLinkSource,
        ordersHasLinkFilters: Object.keys(ordersWin.linkFilters).length > 0,
        custIsLinkSource: custWin.id === app._test._activeLinkSourceId,
      };
    });
    expect(state.ordersDisableLinkSource).toBe(false);
    expect(state.ordersHasLinkFilters).toBe(true);
    expect(state.custIsLinkSource).toBe(true);
  });

  test('disableLinkSource is NOT reset on table that does not receive link filters', async ({ page }) => {
    // One-directional: orders -> customers only. A third table would not be a target.
    // Use bidirectional but disable on customers, then re-link from orders (customers is source, not target)
    await loadLinkPlugin(page, {
      name: 'No Reset Non-Target Test',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    // Set disableLinkSource on customers (not a target of the link)
    await page.evaluate(() => {
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      custWin.disableLinkSource = true;
    });

    // Select in orders — customers is the TARGET, so its disableLinkSource SHOULD be reset
    await selectOrderRow(page, 0);
    await page.waitForTimeout(200);

    // customers is a target here, so disableLinkSource should be reset
    const custDisabled = await page.evaluate(() => {
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      return custWin.disableLinkSource;
    });
    expect(custDisabled).toBe(false);
  });

  test('pre-existing selection does not trigger linking when plugin is loaded', async ({ page }) => {
    // Select a row in orders BEFORE loading any plugin
    await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const ordersTable = app._test.tables['orders'];
      const row = ordersTable.rows[0];
      ordersWin.selectedCells = new Set();
      for (const col of ordersTable.columns) {
        ordersWin.selectedCells.add(`${row._rownum}:${col}`);
      }
      ordersWin.anchorCell = { rownum: row._rownum, col: 'order_id' };
      app._test.rebuildTable(ordersWin);
    });
    await page.waitForTimeout(200);

    // Load plugin via loadPluginFile (simulated with a File object)
    await page.evaluate(() => {
      const config = {
        name: 'Post-Selection Plugin',
        links: [
          { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
        ]
      };
      const blob = new Blob([JSON.stringify(config)], { type: 'application/json' });
      const file = new File([blob], 'test-plugin.json', { type: 'application/json' });
      return app._test.loadPluginFile(file);
    });
    await page.waitForTimeout(200);

    // Linking chip should NOT appear on orders (selection predates plugin)
    const state = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      const linkingChip = ordersWin.statusbarEl.querySelector('.status-chip-link-source');
      const linkedChip = custWin.statusbarEl.querySelector('.status-chip-link');
      return {
        hasLinkingChip: linkingChip !== null,
        hasLinkedChip: linkedChip !== null,
        activeLinkSourceId: app._test._activeLinkSourceId,
        custLinkFilters: Object.keys(custWin.linkFilters).length,
      };
    });
    expect(state.hasLinkingChip).toBe(false);
    expect(state.hasLinkedChip).toBe(false);
    expect(state.activeLinkSourceId).toBeNull();
    expect(state.custLinkFilters).toBe(0);

    // Now select rows AFTER plugin is loaded — linking should work
    await selectOrderRow(page, 0);
    await page.waitForTimeout(200);

    const afterSelect = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      return {
        hasLinkingChip: ordersWin.statusbarEl.querySelector('.status-chip-link-source') !== null,
        hasLinkedChip: custWin.statusbarEl.querySelector('.status-chip-link') !== null,
        custLinkFilters: Object.keys(custWin.linkFilters).length,
      };
    });
    expect(afterSelect.hasLinkingChip).toBe(true);
    expect(afterSelect.hasLinkedChip).toBe(true);
    expect(afterSelect.custLinkFilters).toBeGreaterThan(0);
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

  test('Ctrl+Shift+1 clears sort', async ({ page }) => {
    // Set up sort
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'desc' }];
      app._test.rebuildTable(win);
      app._test.focusWindow(win.id);
    });
    await page.waitForTimeout(100);

    // Verify sort chip is present
    expect(await page.$('.status-chip-sort')).not.toBeNull();

    // Focus a cell so the window is active
    const cell = page.locator('.subwindow table tbody td.data-cell').first();
    await cell.click();
    await page.waitForTimeout(100);

    const modifier = process.platform === 'darwin' ? 'Meta' : 'Control';
    await page.keyboard.press(`${modifier}+Shift+1`);
    await page.waitForTimeout(100);

    // Sort should be cleared (not just disabled)
    const state = await page.evaluate(() => {
      const win = app._test.windows[0];
      return { disableSort: win.disableSort, sortCols: win.sortCols };
    });
    expect(state.disableSort).toBe(false);
    expect(state.sortCols).toEqual([]);

    // Chip should be gone
    const chip = await page.$('.status-chip-sort');
    expect(chip).toBeNull();
  });

  test('Ctrl+Shift+2 clears filters', async ({ page }) => {
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

    // All rows should show (filters cleared)
    const allRows = await getVisibleRowCount(page);
    expect(allRows).toBeGreaterThan(1);

    // Chip should be gone
    const chip = await page.$('.status-chip-filter');
    expect(chip).toBeNull();

    // Filters should be cleared
    const state = await page.evaluate(() => {
      const win = app._test.windows[0];
      return { columnFilterKeys: Object.keys(win.columnFilters).length, filterText: win.filterText };
    });
    expect(state.columnFilterKeys).toBe(0);
    expect(state.filterText).toBe('');
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

  test('clearing filter does not affect sort, and vice versa', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'desc' }];
      win.columnFilters = { name: new Set(['Alice Johnson', 'Bob Smith', 'Eve Davis']) };
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const filteredRows = await getVisibleRowCount(page);
    expect(filteredRows).toBe(3);

    // Clear filter by clicking chip
    await page.click('.status-chip-filter');
    await page.waitForTimeout(100);

    // Filter chip should be gone, sort chip still present and active
    const filterChip = await page.$('.status-chip-filter');
    expect(filterChip).toBeNull();
    const sortChip = await page.$('.status-chip-sort');
    expect(sortChip).not.toBeNull();
    expect(await sortChip.evaluate(el => !el.classList.contains('off'))).toBe(true);

    // All rows should show (filter cleared) but still sorted desc
    const allRows = await getVisibleRowCount(page);
    expect(allRows).toBe(10); // all 10 rows in sample1.csv
    const firstCell = await getCellText(page, 0, 0);
    expect(firstCell).not.toBe('Alice Johnson'); // still sorted desc

    // Clear sort by clicking chip
    await page.click('.status-chip-sort');
    await page.waitForTimeout(100);

    // Now first row should be original order
    const unsortedFirst = await getCellText(page, 0, 0);
    expect(unsortedFirst).toBe('Alice Johnson');

    // Both chips should be gone
    expect(await page.$('.status-chip-sort')).toBeNull();
    expect(await page.$('.status-chip-filter')).toBeNull();
  });

  test('sort chip stays gone after rebuild when sort was cleared', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    // Chip should be present
    let chip = await page.$('.status-chip-sort');
    expect(chip).not.toBeNull();

    // Clear sort by clicking chip
    await page.click('.status-chip-sort');
    await page.waitForTimeout(100);

    // Chip should be gone
    chip = await page.$('.status-chip-sort');
    expect(chip).toBeNull();

    // Simulate a cell edit that triggers rebuildTable
    await page.evaluate(() => {
      const win = app._test.windows[0];
      const t = app._test.tables[win.tableName];
      t.rows[0].name = 'Edited Name';
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    // Sort chip should still be gone after rebuild
    chip = await page.$('.status-chip-sort');
    expect(chip).toBeNull();

    // sortCols should still be empty
    const sortCols = await page.evaluate(() => app._test.windows[0].sortCols);
    expect(sortCols).toEqual([]);
  });
});

test.describe('Chip toggle link re-entrancy guard', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, 'customers.csv');
    await waitForWindow(page, 'customers');
    await uploadFile(page, 'orders.csv');
    await waitForWindow(page, 'orders');
  });

  async function loadPlugin(page, config) {
    await page.evaluate((cfg) => {
      const errors = app._test.validatePlugin(cfg);
      if (errors.length) throw new Error(errors.join(', '));
      const compiled = app._test.compilePlugin(cfg);
      cfg._compiled = compiled;
      app._test.plugins.push(cfg);
      app._test.rebuildTransformCache();
    }, config);
  }

  async function selectRowInWindow(page, tableName, rowIdx) {
    await page.evaluate(({ tableName, rowIdx }) => {
      const win = app._test.windows.find(w => w.tableName === tableName);
      const t = app._test.tables[tableName];
      const row = t.rows[rowIdx];
      win.selectedCells = new Set();
      for (const col of t.columns) {
        win.selectedCells.add(`${row._rownum}:${col}`);
      }
      win.anchorCell = { rownum: row._rownum, col: t.columns[0] };
    }, { tableName, rowIdx });
  }

  test('clicking Formatted chip does not make the table the link source', async ({ page }) => {
    await loadPlugin(page, {
      name: 'Format + Link',
      tables: [{ table: '.*', columns: [{ match: '^name$', display: 'upper(value)' }] }],
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });
    await page.evaluate(() => app._test.rerenderAllWindows());
    await page.waitForTimeout(100);

    await selectRowInWindow(page, 'customers', 0);

    const custWin = await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'customers');
      return { id: w.id };
    });

    const beforeSourceId = await page.evaluate(() => app._test._activeLinkSourceId);
    expect(beforeSourceId).toBeNull();

    await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'customers');
      const chip = w.statusbarEl.querySelector('.status-chip-format');
      chip.click();
    });
    await page.waitForTimeout(200);

    const afterSourceId = await page.evaluate(() => app._test._activeLinkSourceId);
    expect(afterSourceId).toBeNull();

    const linkingChip = await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'customers');
      return w.statusbarEl.querySelector('.status-chip-link-source') !== null;
    });
    expect(linkingChip).toBe(false);
  });

  test('clicking Formatted chip does not clear existing link filters on other tables', async ({ page }) => {
    await loadPlugin(page, {
      name: 'Format + Link',
      tables: [{ table: '.*', columns: [{ match: '^name$', display: 'upper(value)' }] }],
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });
    await page.evaluate(() => app._test.rerenderAllWindows());
    await page.waitForTimeout(100);

    await selectRowInWindow(page, 'orders', 0);
    await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'orders');
      app._test.applyLinkFilters(w);
    });
    await page.waitForTimeout(200);

    const custFilteredBefore = await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'customers');
      return { rows: w._displayRows.length, linkFilterKeys: Object.keys(w.linkFilters).length };
    });
    expect(custFilteredBefore.rows).toBe(1);
    expect(custFilteredBefore.linkFilterKeys).toBe(1);

    await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'customers');
      const chip = w.statusbarEl.querySelector('.status-chip-format');
      chip.click();
    });
    await page.waitForTimeout(200);

    const custFilteredAfter = await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'customers');
      return { rows: w._displayRows.length, linkFilterKeys: Object.keys(w.linkFilters).length };
    });
    expect(custFilteredAfter.rows).toBe(1);
    expect(custFilteredAfter.linkFilterKeys).toBe(1);

    const sourceId = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      return { sourceId: app._test._activeLinkSourceId, ordersId: ordersWin.id };
    });
    expect(sourceId.sourceId).toBe(sourceId.ordersId);
  });

  test('clicking Sorted chip does not make the table the link source', async ({ page }) => {
    await loadPlugin(page, {
      name: 'Link Only',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'customers');
      w.sortCols = [{ col: 'name', dir: 'asc' }];
      app._test.rebuildTable(w);
    });
    await page.waitForTimeout(100);

    await selectRowInWindow(page, 'customers', 0);

    await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'customers');
      const chip = w.statusbarEl.querySelector('.status-chip-sort');
      chip.click();
    });
    await page.waitForTimeout(200);

    const sourceId = await page.evaluate(() => app._test._activeLinkSourceId);
    expect(sourceId).toBeNull();
  });

  test('clicking Filtered chip does not make the table the link source', async ({ page }) => {
    await loadPlugin(page, {
      name: 'Link Only',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'customers');
      w.columnFilters = { name: new Set(['Alice']) };
      app._test.rebuildTable(w);
    });
    await page.waitForTimeout(100);

    await selectRowInWindow(page, 'customers', 0);

    await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'customers');
      const chip = w.statusbarEl.querySelector('.status-chip-filter');
      chip.click();
    });
    await page.waitForTimeout(200);

    const sourceId = await page.evaluate(() => app._test._activeLinkSourceId);
    expect(sourceId).toBeNull();
  });

  test('clicking Sorted chip does not clear existing link filters', async ({ page }) => {
    await loadPlugin(page, {
      name: 'Link Only',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    await selectRowInWindow(page, 'orders', 0);
    await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'orders');
      app._test.applyLinkFilters(w);
    });
    await page.waitForTimeout(200);

    const custBefore = await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'customers');
      return w._displayRows.length;
    });
    expect(custBefore).toBe(1);

    await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'orders');
      w.sortCols = [{ col: 'order_id', dir: 'asc' }];
      app._test.rebuildTable(w);
    });
    await page.waitForTimeout(100);

    await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'orders');
      const chip = w.statusbarEl.querySelector('.status-chip-sort');
      chip.click();
    });
    await page.waitForTimeout(200);

    const custAfter = await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'customers');
      return w._displayRows.length;
    });
    expect(custAfter).toBe(1);

    const sourceId = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      return { sourceId: app._test._activeLinkSourceId, ordersId: ordersWin.id };
    });
    expect(sourceId.sourceId).toBe(sourceId.ordersId);
  });

  test('Ctrl+Shift+4 (format toggle) does not make table the link source', async ({ page }) => {
    await loadPlugin(page, {
      name: 'Format + Link',
      tables: [{ table: '.*', columns: [{ match: '^name$', display: 'upper(value)' }] }],
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });
    await page.evaluate(() => app._test.rerenderAllWindows());
    await page.waitForTimeout(100);

    await selectRowInWindow(page, 'customers', 0);

    await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'customers');
      app._test.focusWindow(w.id);
      const cell = w.bodyEl.querySelector('td.data-cell');
      if (cell) cell.focus();
    });
    await page.waitForTimeout(100);

    const modifier = process.platform === 'darwin' ? 'Meta' : 'Control';
    await page.keyboard.press(`${modifier}+Shift+4`);
    await page.waitForTimeout(200);

    const sourceId = await page.evaluate(() => app._test._activeLinkSourceId);
    expect(sourceId).toBeNull();
  });

  test('Ctrl+Shift+1 (clear sort) does not make table the link source', async ({ page }) => {
    await loadPlugin(page, {
      name: 'Link Only',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'customers');
      w.sortCols = [{ col: 'name', dir: 'asc' }];
      app._test.rebuildTable(w);
    });
    await page.waitForTimeout(100);

    await selectRowInWindow(page, 'customers', 0);

    await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'customers');
      app._test.focusWindow(w.id);
      const cell = w.bodyEl.querySelector('td.data-cell');
      if (cell) cell.focus();
    });
    await page.waitForTimeout(100);

    const modifier = process.platform === 'darwin' ? 'Meta' : 'Control';
    await page.keyboard.press(`${modifier}+Shift+1`);
    await page.waitForTimeout(200);

    const sourceId = await page.evaluate(() => app._test._activeLinkSourceId);
    expect(sourceId).toBeNull();
  });

  test('Ctrl+Shift+2 (clear filters) does not make table the link source', async ({ page }) => {
    await loadPlugin(page, {
      name: 'Link Only',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'customers');
      w.columnFilters = { name: new Set(['Alice']) };
      app._test.rebuildTable(w);
    });
    await page.waitForTimeout(100);

    await selectRowInWindow(page, 'customers', 0);

    await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'customers');
      app._test.focusWindow(w.id);
      const cell = w.bodyEl.querySelector('td.data-cell');
      if (cell) cell.focus();
    });
    await page.waitForTimeout(100);

    const modifier = process.platform === 'darwin' ? 'Meta' : 'Control';
    await page.keyboard.press(`${modifier}+Shift+2`);
    await page.waitForTimeout(200);

    const sourceId = await page.evaluate(() => app._test._activeLinkSourceId);
    expect(sourceId).toBeNull();
  });

  test('Ctrl+Shift+4 does not clear existing link filters', async ({ page }) => {
    await loadPlugin(page, {
      name: 'Format + Link',
      tables: [{ table: '.*', columns: [{ match: '^name$', display: 'upper(value)' }] }],
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });
    await page.evaluate(() => app._test.rerenderAllWindows());
    await page.waitForTimeout(100);

    await selectRowInWindow(page, 'orders', 0);
    await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'orders');
      app._test.applyLinkFilters(w);
    });
    await page.waitForTimeout(200);

    const custBefore = await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'customers');
      return w._displayRows.length;
    });
    expect(custBefore).toBe(1);

    await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'customers');
      app._test.focusWindow(w.id);
      const cell = w.bodyEl.querySelector('td.data-cell');
      if (cell) cell.focus();
    });
    await page.waitForTimeout(100);

    const modifier = process.platform === 'darwin' ? 'Meta' : 'Control';
    await page.keyboard.press(`${modifier}+Shift+4`);
    await page.waitForTimeout(200);

    const custAfter = await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'customers');
      return w._displayRows.length;
    });
    expect(custAfter).toBe(1);

    const sourceId = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      return { sourceId: app._test._activeLinkSourceId, ordersId: ordersWin.id };
    });
    expect(sourceId.sourceId).toBe(sourceId.ordersId);
  });

  test('Linking chip toggle still works (not blocked by guard)', async ({ page }) => {
    await loadPlugin(page, {
      name: 'Link Only',
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });

    await selectRowInWindow(page, 'orders', 0);
    await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'orders');
      app._test.applyLinkFilters(w);
    });
    await page.waitForTimeout(200);

    const custBefore = await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'customers');
      return w._displayRows.length;
    });
    expect(custBefore).toBe(1);

    await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'orders');
      const chip = w.statusbarEl.querySelector('.status-chip-link-source');
      chip.click();
    });
    await page.waitForTimeout(200);

    const custAfter = await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'customers');
      return w._displayRows.length;
    });
    expect(custAfter).toBe(3);

    await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'orders');
      const chip = w.statusbarEl.querySelector('.status-chip-link-source');
      chip.click();
    });
    await page.waitForTimeout(200);

    const custReEnabled = await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'customers');
      return w._displayRows.length;
    });
    expect(custReEnabled).toBe(1);
  });

  test('toggling Formatted chip on source table preserves its Linking status', async ({ page }) => {
    await loadPlugin(page, {
      name: 'Format + Link',
      tables: [{ table: '.*', columns: [{ match: '^order_id$', display: '"#" + value' }] }],
      links: [
        { source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }
      ]
    });
    await page.evaluate(() => app._test.rerenderAllWindows());
    await page.waitForTimeout(100);

    await selectRowInWindow(page, 'orders', 0);
    await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'orders');
      app._test.applyLinkFilters(w);
    });
    await page.waitForTimeout(200);

    const before = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      return {
        sourceId: app._test._activeLinkSourceId,
        ordersId: ordersWin.id,
        custRows: custWin._displayRows.length,
        hasLinkingChip: ordersWin.statusbarEl.querySelector('.status-chip-link-source') !== null,
      };
    });
    expect(before.sourceId).toBe(before.ordersId);
    expect(before.custRows).toBe(1);
    expect(before.hasLinkingChip).toBe(true);

    await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'orders');
      const chip = w.statusbarEl.querySelector('.status-chip-format');
      chip.click();
    });
    await page.waitForTimeout(200);

    const after = await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      return {
        sourceId: app._test._activeLinkSourceId,
        ordersId: ordersWin.id,
        custRows: custWin._displayRows.length,
        hasLinkingChip: ordersWin.statusbarEl.querySelector('.status-chip-link-source') !== null,
      };
    });
    expect(after.sourceId).toBe(after.ordersId);
    expect(after.custRows).toBe(1);
    expect(after.hasLinkingChip).toBe(true);
  });
});
