const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow } = require('../helpers');
const path = require('path');

test.describe('Column Header Badges', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('no badges shown on unsorted, unfiltered columns', async ({ page }) => {
    const badges = await page.$$('.col-badges');
    expect(badges.length).toBe(0);
  });

  test('sort badge appears on ascending sort', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const badge = await page.$('th.sorted .col-badge-sort.sort-asc');
    expect(badge).not.toBeNull();
  });

  test('sort badge flips direction on descending sort', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'desc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const badge = await page.$('th.sorted .col-badge-sort.sort-desc');
    expect(badge).not.toBeNull();
    const ascBadge = await page.$('th.sorted .col-badge-sort.sort-asc');
    expect(ascBadge).toBeNull();
  });

  test('multi-sort badges show numbers inside triangles', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }, { col: 'email', dir: 'desc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const nums = await page.$$eval('.sort-num', els => els.map(e => e.textContent));
    expect(nums).toEqual(['1', '2']);
  });

  test('single sort does not show number inside triangle', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const nums = await page.$$('.sort-num');
    expect(nums.length).toBe(0);
  });

  test('filter badge appears when column autofilter is applied', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.columnFilters['name'] = new Set(['Alice Johnson']);
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const badge = await page.$('th.col-filtered .col-badge-filter');
    expect(badge).not.toBeNull();
    const vis = await badge.evaluate(el => getComputedStyle(el).visibility);
    expect(vis).toBe('visible');
  });

  test('filter badge disappears when filter is cleared', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.columnFilters['name'] = new Set(['Alice Johnson']);
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.columnFilters = {};
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const badges = await page.$$('.col-badge-filter');
    expect(badges.length).toBe(0);
  });

  test('sort and filter badges coexist on the same column', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }];
      win.columnFilters['name'] = new Set(['Alice Johnson']);
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const th = page.locator('th.sorted.col-filtered');
    const filterBadge = th.locator('.col-badge-filter');
    const sortBadge = th.locator('.col-badge-sort');
    expect(await filterBadge.count()).toBe(1);
    expect(await sortBadge.count()).toBe(1);
  });

  test('badge order is link, format, filter, sort', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }];
      win.columnFilters['name'] = new Set(['Alice Johnson']);
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const classes = await page.$$eval(
      'th.sorted .col-badge',
      els => els.map(el => {
        if (el.classList.contains('col-badge-link')) return 'link';
        if (el.classList.contains('col-badge-format')) return 'format';
        if (el.classList.contains('col-badge-filter')) return 'filter';
        if (el.classList.contains('col-badge-sort')) return 'sort';
        return 'unknown';
      })
    );
    expect(classes).toEqual(['link', 'format', 'filter', 'sort']);
  });

  test('all four badge slots are rendered when any badge is active', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const badges = await page.$$eval(
      'th.sorted .col-badge',
      els => els.map(el => ({
        type: el.classList.contains('col-badge-sort') ? 'sort' :
              el.classList.contains('col-badge-filter') ? 'filter' :
              el.classList.contains('col-badge-link') ? 'link' : 'format',
        visible: getComputedStyle(el).visibility === 'visible'
      }))
    );
    expect(badges).toEqual([
      { type: 'link', visible: false },
      { type: 'format', visible: false },
      { type: 'filter', visible: false },
      { type: 'sort', visible: true },
    ]);
  });

  test('badges stay in fixed positions when other badges toggle', async ({ page }) => {
    // Sort only — record sort badge x position
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const sortX1 = await page.$eval('th.sorted .col-badge-sort',
      el => el.getBoundingClientRect().x);

    // Add filter — sort badge should not move
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.columnFilters['name'] = new Set(['Alice Johnson']);
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const sortX2 = await page.$eval('th.sorted .col-badge-sort',
      el => el.getBoundingClientRect().x);

    expect(Math.abs(sortX1 - sortX2)).toBeLessThan(1);
  });

  test('badges are positioned at top-right of header cell', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const pos = await page.evaluate(() => {
      const th = document.querySelector('th.sorted');
      const badges = th.querySelector('.col-badges');
      const thRect = th.getBoundingClientRect();
      const bRect = badges.getBoundingClientRect();
      return {
        nearTop: bRect.y - thRect.y < 5,
        nearRight: (thRect.x + thRect.width) - (bRect.x + bRect.width) < 25,
      };
    });

    expect(pos.nearTop).toBe(true);
    expect(pos.nearRight).toBe(true);
  });

  test('badges overlap the filter button', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const overlap = await page.evaluate(() => {
      const th = document.querySelector('th.sorted');
      const badges = th.querySelector('.col-badges');
      const filterBtn = th.querySelector('.col-filter-btn');
      const bRect = badges.getBoundingClientRect();
      const fRect = filterBtn.getBoundingClientRect();
      return (bRect.x + bRect.width) - fRect.x;
    });

    expect(overlap).toBeGreaterThan(0);
  });

  test('no old sort-arrow elements exist', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const arrows = await page.$$('.sort-arrow');
    expect(arrows.length).toBe(0);
  });

  test('no colored left-border on sorted column header', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const borderLeft = await page.$eval('th.sorted',
      el => getComputedStyle(el).borderLeftWidth);
    expect(borderLeft).toBe('1px');
  });

  test('no bottom border on sorted column header', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const borderBottom = await page.$eval('th.sorted',
      el => getComputedStyle(el).borderBottomWidth);
    expect(borderBottom).toBe('1px');
  });

  test('no colored left-border on filtered column header', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.columnFilters['name'] = new Set(['Alice Johnson']);
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const borderLeft = await page.$eval('th.col-filtered',
      el => getComputedStyle(el).borderLeftWidth);
    expect(borderLeft).toBe('1px');
  });

  test('sort number color is white for contrast', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }, { col: 'email', dir: 'desc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const color = await page.$eval('.sort-num',
      el => getComputedStyle(el).color);
    expect(color).toBe('rgb(255, 255, 255)');
  });

  test('badges have pointer-events none so they do not block clicks', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const pe = await page.$eval('.col-badges',
      el => getComputedStyle(el).pointerEvents);
    expect(pe).toBe('none');
  });

  test('suspended format badge is hidden but takes space', async ({ page }) => {
    // Load a plugin with a transform
    await page.evaluate(() => {
      const cfg = {
        name: 'Test Transform',
        table: '.*',
        columns: [{ match: '^name$', display: '"[" + value + "]"' }]
      };
      const compiled = app._test.compilePlugin(cfg);
      cfg._compiled = compiled;
      app._test.plugins.push(cfg);
      app._test.rebuildTransformCache();
      const win = app._test.windows[0];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(200);

    // Verify format badge is visible
    let vis = await page.evaluate(() => {
      const th = document.querySelector('th.col-transformed');
      if (!th) return null;
      const badge = th.querySelector('.col-badge-format');
      return badge ? getComputedStyle(badge).visibility : null;
    });
    expect(vis).toBe('visible');

    // Suspend formatting
    await page.evaluate(() => {
      const win = app._test.windows[0];
      const t = app._test.tables[win.tableName];
      const cols = t.columns.filter(c => app._test.hasDisplayTransform(win.tableName, c));
      cols.forEach(c => win.disabledTransforms.add(c));
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    // Format badge should be hidden
    vis = await page.evaluate(() => {
      const th = document.querySelector('th.col-transformed');
      if (!th) return null;
      const badge = th.querySelector('.col-badge-format');
      return badge ? getComputedStyle(badge).visibility : null;
    });
    expect(vis).toBe('hidden');

    // Unload plugin
    await page.evaluate(() => app._test.unloadPlugin(0));
  });

  test('chip order matches badge order: Linked, Linking, Formatted, Filtered, Sorted', async ({ page }) => {
    // Apply sort + filter to get both chips visible
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }];
      win.columnFilters['name'] = new Set(['Alice Johnson']);
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const chipTexts = await page.$$eval('.status-chip', els => els.map(e => e.textContent));
    expect(chipTexts).toEqual(['Filtered', 'Sorted']);

    // Badge order should be: link, format, filter, sort
    const badgeTypes = await page.$$eval('th.sorted .col-badge', els => els.map(el => {
      if (el.classList.contains('col-badge-link')) return 'link';
      if (el.classList.contains('col-badge-format')) return 'format';
      if (el.classList.contains('col-badge-filter')) return 'filter';
      if (el.classList.contains('col-badge-sort')) return 'sort';
      return 'unknown';
    }));
    expect(badgeTypes).toEqual(['link', 'format', 'filter', 'sort']);
  });
});

test.describe('Column Badge Link Isolation', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/customers.csv');
    await waitForWindow(page, 'customers');
    await uploadFile(page, '../test/orders.csv');
    await waitForWindow(page, 'orders');
  });

  test('sorting a linked target table does not clear its link filters', async ({ page }) => {
    // Load link plugin
    await page.evaluate(() => {
      const cfg = {
        name: 'Link Test',
        links: [{ source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }]
      };
      const compiled = app._test.compilePlugin(cfg);
      cfg._compiled = compiled;
      app._test.plugins.push(cfg);
      app._test.rebuildTransformCache();
    });
    await page.waitForTimeout(200);

    // Select a row in orders to establish link
    await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const row = ordersWin._displayRows[0];
      ordersWin.selectedCells = new Set([`${row._rownum}:customer_id`]);
      ordersWin.anchorCell = { rownum: row._rownum, col: 'customer_id' };
      app._test.applyLinkFilters(ordersWin);
    });
    await page.waitForTimeout(200);

    // Verify customers table is filtered by link
    const custRowsBefore = await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'customers');
      return w._displayRows.length;
    });
    expect(custRowsBefore).toBeLessThan(3);

    // Now sort the customers table (target) — link filters should survive
    await page.evaluate(() => {
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      custWin.sortCols = [{ col: 'name', dir: 'asc' }];
      app._test.rebuildTable(custWin);
    });
    await page.waitForTimeout(200);

    const custRowsAfter = await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'customers');
      return w._displayRows.length;
    });
    expect(custRowsAfter).toBe(custRowsBefore);

    // Linked chip should still be present
    const hasLinkedChip = await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'customers');
      return !!w.statusbarEl.querySelector('.status-chip-link');
    });
    expect(hasLinkedChip).toBe(true);

    await page.evaluate(() => app._test.unloadPlugin(0));
  });

  test('filtering a linked target table does not clear its link filters', async ({ page }) => {
    await page.evaluate(() => {
      const cfg = {
        name: 'Link Test',
        links: [{ source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }]
      };
      const compiled = app._test.compilePlugin(cfg);
      cfg._compiled = compiled;
      app._test.plugins.push(cfg);
      app._test.rebuildTransformCache();
    });
    await page.waitForTimeout(200);

    await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      const row = ordersWin._displayRows[0];
      ordersWin.selectedCells = new Set([`${row._rownum}:customer_id`]);
      ordersWin.anchorCell = { rownum: row._rownum, col: 'customer_id' };
      app._test.applyLinkFilters(ordersWin);
    });
    await page.waitForTimeout(200);

    const custRowsBefore = await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'customers');
      return Object.keys(w.linkFilters).length;
    });
    expect(custRowsBefore).toBeGreaterThan(0);

    // Apply a column autofilter on the target table
    await page.evaluate(() => {
      const custWin = app._test.windows.find(w => w.tableName === 'customers');
      custWin.columnFilters['name'] = new Set(['Alice']);
      app._test.rebuildTable(custWin);
    });
    await page.waitForTimeout(200);

    const linkFiltersAfter = await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'customers');
      return Object.keys(w.linkFilters).length;
    });
    expect(linkFiltersAfter).toBe(custRowsBefore);

    await page.evaluate(() => app._test.unloadPlugin(0));
  });

  test('sorting a source table does not re-engage linking when no selection exists', async ({ page }) => {
    await page.evaluate(() => {
      const cfg = {
        name: 'Link Test',
        links: [{ source: { table: 'orders', column: 'customer_id' }, target: { table: 'customers', column: 'id' } }]
      };
      const compiled = app._test.compilePlugin(cfg);
      cfg._compiled = compiled;
      app._test.plugins.push(cfg);
      app._test.rebuildTransformCache();
    });
    await page.waitForTimeout(200);

    // Sort orders without selecting any cells — should not create link filters
    await page.evaluate(() => {
      const ordersWin = app._test.windows.find(w => w.tableName === 'orders');
      ordersWin.sortCols = [{ col: 'customer_id', dir: 'asc' }];
      app._test.rebuildTable(ordersWin);
    });
    await page.waitForTimeout(200);

    const custLinkFilters = await page.evaluate(() => {
      const w = app._test.windows.find(w => w.tableName === 'customers');
      return Object.keys(w.linkFilters).length;
    });
    expect(custLinkFilters).toBe(0);

    await page.evaluate(() => app._test.unloadPlugin(0));
  });
});
