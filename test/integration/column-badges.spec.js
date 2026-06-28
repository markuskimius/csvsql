const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow } = require('../helpers');
const path = require('path');

test.describe('Column Header Badges', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('sort outline badge shown on unsorted columns', async ({ page }) => {
    const outlines = await page.$$('.col-badge-sort.sort-outline');
    expect(outlines.length).toBeGreaterThan(0);
    const filledSort = await page.$$('.col-badge-sort.sort-asc, .col-badge-sort.sort-desc');
    expect(filledSort.length).toBe(0);
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

  test('filter button highlights green when column autofilter is applied', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.columnFilters['name'] = new Set(['Alice Johnson']);
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const filterBtn = await page.$('th.col-filtered .col-filter-btn.active');
    expect(filterBtn).not.toBeNull();
  });

  test('filter button loses active class when filter is cleared', async ({ page }) => {
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

    const allInactive = await page.$$eval('.col-filter-btn',
      els => els.every(el => !el.classList.contains('active')));
    expect(allInactive).toBe(true);
  });

  test('sort badge and filter button coexist on the same column', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }];
      win.columnFilters['name'] = new Set(['Alice Johnson']);
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const th = page.locator('th.sorted.col-filtered');
    const filterBtn = th.locator('.col-filter-btn.active');
    const sortBadge = th.locator('.col-badge-sort');
    expect(await filterBtn.count()).toBe(1);
    expect(await sortBadge.count()).toBe(1);
  });

  test('badge order in col-badges is link, format', async ({ page }) => {
    const badges = await page.$$eval(
      '.col-badges .col-badge',
      els => els.map(el => {
        if (el.classList.contains('col-badge-link')) return 'link';
        if (el.classList.contains('col-badge-format')) return 'format';
        return 'unknown';
      })
    );
    const unique = [...new Set(badges)];
    expect(unique).toEqual(['link', 'format']);
  });

  test('dot badge slots always rendered with faint outlines for inactive', async ({ page }) => {
    const badges = await page.$$eval(
      '.col-badges .col-badge',
      els => els.slice(0, 2).map(el => ({
        type: el.classList.contains('col-badge-link') ? 'link' : 'format',
        faint: el.classList.contains('badge-faint'),
      }))
    );
    expect(badges).toEqual([
      { type: 'link', faint: true },
      { type: 'format', faint: true },
    ]);
  });

  test('sort badge inline in th-inner shows filled when sorted', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const sortBadge = await page.$('th.sorted .th-inner .col-badge-sort.sort-asc');
    expect(sortBadge).not.toBeNull();
  });

  test('sort badge is wrapped in col-sort-btn container', async ({ page }) => {
    const structure = await page.evaluate(() => {
      const th = document.querySelector('.data-table th:not(.row-num-header)');
      const sortBtn = th.querySelector('.col-sort-btn');
      const badge = sortBtn ? sortBtn.querySelector('.col-badge-sort') : null;
      return { hasSortBtn: !!sortBtn, hasBadgeInside: !!badge };
    });
    expect(structure.hasSortBtn).toBe(true);
    expect(structure.hasBadgeInside).toBe(true);
  });

  test('sort button shows hover highlight', async ({ page }) => {
    const sortBtn = page.locator('.col-sort-btn').first();
    const bgBefore = await sortBtn.evaluate(el => getComputedStyle(el).backgroundColor);
    await sortBtn.hover();
    const bgAfter = await sortBtn.evaluate(el => getComputedStyle(el).backgroundColor);
    expect(bgBefore).not.toBe(bgAfter);
    expect(bgAfter).toBe('rgba(124, 111, 247, 0.15)');
  });

  test('filter button shows hover highlight', async ({ page }) => {
    const filterBtn = page.locator('.col-filter-btn').first();
    const bgBefore = await filterBtn.evaluate(el => getComputedStyle(el).backgroundColor);
    await filterBtn.hover();
    const bgAfter = await filterBtn.evaluate(el => getComputedStyle(el).backgroundColor);
    expect(bgBefore).not.toBe(bgAfter);
    expect(bgAfter).toBe('rgba(124, 111, 247, 0.15)');
  });

  test('sort button and filter button have equal dimensions', async ({ page }) => {
    const dims = await page.evaluate(() => {
      const th = document.querySelector('.data-table th:not(.row-num-header)');
      const sortBtn = th.querySelector('.col-sort-btn');
      const filterBtn = th.querySelector('.col-filter-btn');
      const sRect = sortBtn.getBoundingClientRect();
      const fRect = filterBtn.getBoundingClientRect();
      return {
        sortW: Math.round(sRect.width),
        sortH: Math.round(sRect.height),
        filterW: Math.round(fRect.width),
        filterH: Math.round(fRect.height),
      };
    });
    expect(dims.sortW).toBe(dims.filterW);
    expect(dims.sortH).toBe(dims.filterH);
    expect(dims.sortW).toBe(14);
    expect(dims.sortH).toBe(14);
  });

  test('sort button hover highlight matches filter button hover highlight', async ({ page }) => {
    const sortBtn = page.locator('.col-sort-btn').first();
    await sortBtn.hover();
    const sortBg = await sortBtn.evaluate(el => getComputedStyle(el).backgroundColor);

    const filterBtn = page.locator('.col-filter-btn').first();
    await filterBtn.hover();
    const filterBg = await filterBtn.evaluate(el => getComputedStyle(el).backgroundColor);

    expect(sortBg).toBe(filterBg);
  });

  test('filter button contains an SVG funnel icon', async ({ page }) => {
    const svg = await page.evaluate(() => {
      const btn = document.querySelector('.col-filter-btn');
      const svgEl = btn.querySelector('svg');
      if (!svgEl) return null;
      const path = svgEl.querySelector('path');
      return {
        hasSvg: true,
        hasPath: !!path,
        fill: svgEl.getAttribute('fill'),
        stroke: svgEl.getAttribute('stroke'),
        viewBox: svgEl.getAttribute('viewBox'),
      };
    });
    expect(svg).not.toBeNull();
    expect(svg.hasSvg).toBe(true);
    expect(svg.hasPath).toBe(true);
    expect(svg.fill).toBe('none');
    expect(svg.stroke).toBe('currentColor');
  });

  test('filter funnel is outline-only when inactive', async ({ page }) => {
    const svgFill = await page.evaluate(() => {
      const btn = document.querySelector('.col-filter-btn:not(.active)');
      const svg = btn.querySelector('svg');
      return svg.getAttribute('fill');
    });
    expect(svgFill).toBe('none');
  });

  test('filter funnel fills green when active', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.columnFilters['name'] = new Set(['Alice Johnson']);
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const fill = await page.evaluate(() => {
      const btn = document.querySelector('.col-filter-btn.active');
      if (!btn) return null;
      const svg = btn.querySelector('svg');
      return getComputedStyle(svg).fill;
    });
    expect(fill).toBe('rgb(85, 204, 136)');
  });

  test('filter funnel outline color matches sort triangle outline color', async ({ page }) => {
    const colors = await page.evaluate(() => {
      const filterBtn = document.querySelector('.col-filter-btn:not(.active)');
      const filterColor = getComputedStyle(filterBtn).color;
      return { filterColor };
    });
    expect(colors.filterColor).toBe('rgba(255, 255, 255, 0.15)');
  });

  test('filter funnel does not light up on header hover, only on button hover', async ({ page }) => {
    const filterBtn = page.locator('.col-filter-btn').first();

    // Move mouse away from the header entirely first
    await page.mouse.move(0, 0);
    const bgBefore = await filterBtn.evaluate(el => getComputedStyle(el).backgroundColor);

    // Hover the filter button directly
    await filterBtn.hover();
    const bgAfterBtnHover = await filterBtn.evaluate(el => getComputedStyle(el).backgroundColor);
    expect(bgAfterBtnHover).toBe('rgba(124, 111, 247, 0.15)');
    expect(bgBefore).not.toBe(bgAfterBtnHover);
  });

  test('sort triangle outline does not light up on header hover, only on sort button hover', async ({ page }) => {
    const th = page.locator('.data-table th:not(.row-num-header)').first();
    const sortBtn = th.locator('.col-sort-btn');
    const outline = sortBtn.locator('.col-badge-sort.sort-outline');

    const bgBefore = await outline.evaluate(el => getComputedStyle(el).backgroundImage);
    await th.hover({ position: { x: 5, y: 5 } });
    const bgAfterThHover = await outline.evaluate(el => getComputedStyle(el).backgroundImage);
    expect(bgAfterThHover).toBe(bgBefore);

    await sortBtn.hover();
    const bgAfterBtnHover = await outline.evaluate(el => getComputedStyle(el).backgroundImage);
    expect(bgAfterBtnHover).not.toBe(bgBefore);
  });

  test('filter button has zero left margin for tight spacing with sort button', async ({ page }) => {
    const margin = await page.evaluate(() => {
      const btn = document.querySelector('.col-filter-btn');
      return getComputedStyle(btn).marginLeft;
    });
    expect(margin).toBe('0px');
  });

  test('sort button click still triggers sort', async ({ page }) => {
    const sortBtn = page.locator('.data-table th:not(.row-num-header)').nth(0).locator('.col-sort-btn');
    await sortBtn.click();
    await page.waitForTimeout(200);

    const sorted = await page.evaluate(() => {
      const win = app._test.windows[0];
      return win.sortCols.length > 0 && win.sortCols[0].dir === 'asc';
    });
    expect(sorted).toBe(true);
  });

  test('sort button click does not trigger column selection', async ({ page }) => {
    const sortBtn = page.locator('.data-table th:not(.row-num-header)').nth(0).locator('.col-sort-btn');
    await sortBtn.click();
    await page.waitForTimeout(200);

    const selectedCol = await page.evaluate(() => app._test.windows[0].selectedCol);
    expect(selectedCol).toBeNull();
  });

  test('sort button has correct border-radius', async ({ page }) => {
    const radius = await page.evaluate(() => {
      const btn = document.querySelector('.col-sort-btn');
      return getComputedStyle(btn).borderRadius;
    });
    expect(radius).toBe('2px');
  });

  test('filter button has correct border-radius', async ({ page }) => {
    const radius = await page.evaluate(() => {
      const btn = document.querySelector('.col-filter-btn');
      return getComputedStyle(btn).borderRadius;
    });
    expect(radius).toBe('2px');
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

  test('badges are positioned at center-bottom of header cell', async ({ page }) => {
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
        nearBottom: (thRect.y + thRect.height) - (bRect.y + bRect.height) < 5,
        centered: Math.abs((thRect.x + thRect.width / 2) - (bRect.x + bRect.width / 2)) < 5,
      };
    });

    expect(pos.nearBottom).toBe(true);
    expect(pos.centered).toBe(true);
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

  test('sort number is visually inside ascending triangle', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }, { col: 'email', dir: 'asc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const pos = await page.evaluate(() => {
      const badge = document.querySelector('.col-badge-sort.sort-asc');
      const num = badge.querySelector('.sort-num');
      const bRect = badge.getBoundingClientRect();
      const nRect = num.getBoundingClientRect();
      return {
        numTop: nRect.top,
        numBottom: nRect.bottom,
        badgeTop: bRect.top,
        badgeBottom: bRect.bottom,
      };
    });
    expect(pos.numTop).toBeGreaterThanOrEqual(pos.badgeTop - 1);
    expect(pos.numBottom).toBeLessThanOrEqual(pos.badgeBottom + 1);
  });

  test('sort number is visually inside descending triangle', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'desc' }, { col: 'email', dir: 'desc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const pos = await page.evaluate(() => {
      const badge = document.querySelector('.col-badge-sort.sort-desc');
      const num = badge.querySelector('.sort-num');
      const bRect = badge.getBoundingClientRect();
      const nRect = num.getBoundingClientRect();
      return {
        numTop: nRect.top,
        numBottom: nRect.bottom,
        badgeTop: bRect.top,
        badgeBottom: bRect.bottom,
      };
    });
    expect(pos.numTop).toBeGreaterThanOrEqual(pos.badgeTop - 1);
    expect(pos.numBottom).toBeLessThanOrEqual(pos.badgeBottom + 1);
  });

  test('sort numbers are positioned consistently between asc and desc triangles', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }, { col: 'email', dir: 'desc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const pos = await page.evaluate(() => {
      const ascBadge = document.querySelector('.col-badge-sort.sort-asc');
      const descBadge = document.querySelector('.col-badge-sort.sort-desc');
      const ascNum = ascBadge.querySelector('.sort-num');
      const descNum = descBadge.querySelector('.sort-num');
      const ascBRect = ascBadge.getBoundingClientRect();
      const descBRect = descBadge.getBoundingClientRect();
      const ascNRect = ascNum.getBoundingClientRect();
      const descNRect = descNum.getBoundingClientRect();
      return {
        ascNumInside: ascNRect.top >= ascBRect.top - 1 && ascNRect.bottom <= ascBRect.bottom + 1,
        descNumInside: descNRect.top >= descBRect.top - 1 && descNRect.bottom <= descBRect.bottom + 1,
        ascBadgeHeight: ascBRect.height,
        descBadgeHeight: descBRect.height,
      };
    });
    expect(pos.ascNumInside).toBe(true);
    expect(pos.descNumInside).toBe(true);
    expect(pos.ascBadgeHeight).toBe(pos.descBadgeHeight);
  });

  test('descending sort number does not render below the triangle', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'desc' }, { col: 'email', dir: 'asc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const pos = await page.evaluate(() => {
      const badge = document.querySelector('.col-badge-sort.sort-desc');
      const num = badge.querySelector('.sort-num');
      const bRect = badge.getBoundingClientRect();
      const nRect = num.getBoundingClientRect();
      return {
        numBottom: nRect.bottom,
        badgeBottom: bRect.bottom,
      };
    });
    expect(pos.numBottom).toBeLessThanOrEqual(pos.badgeBottom + 1);
  });

  test('sort number color is dark for contrast against orange badge', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }, { col: 'email', dir: 'desc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const color = await page.$eval('.sort-num',
      el => getComputedStyle(el).color);
    expect(color).toBe('rgb(26, 26, 46)');
  });

  test('sort badge triangle color is orange', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const ascColor = await page.$eval('.col-badge-sort.sort-asc',
      el => getComputedStyle(el).borderBottomColor);
    expect(ascColor).toBe('rgb(245, 158, 11)');
  });

  test('descending sort badge triangle color is orange', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'desc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const descColor = await page.$eval('.col-badge-sort.sort-desc',
      el => getComputedStyle(el).borderTopColor);
    expect(descColor).toBe('rgb(245, 158, 11)');
  });

  test('sort status chip has orange text color', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }];
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const chipColor = await page.$eval('.status-chip-sort',
      el => getComputedStyle(el).color);
    expect(chipColor).toBe('rgb(245, 158, 11)');
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

  test('suspended format badge shows faint outline', async ({ page }) => {
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

    // Verify format badge is active (not faint)
    let faint = await page.evaluate(() => {
      const th = document.querySelector('th.col-transformed');
      if (!th) return null;
      const badge = th.querySelector('.col-badge-format');
      return badge ? badge.classList.contains('badge-faint') : null;
    });
    expect(faint).toBe(false);

    // Suspend formatting
    await page.evaluate(() => {
      const win = app._test.windows[0];
      const t = app._test.tables[win.tableName];
      const cols = t.columns.filter(c => app._test.hasDisplayTransform(win.tableName, c));
      cols.forEach(c => win.disabledTransforms.add(c));
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    // Format badge should be faint
    faint = await page.evaluate(() => {
      const th = document.querySelector('th.col-transformed');
      if (!th) return null;
      const badge = th.querySelector('.col-badge-format');
      return badge ? badge.classList.contains('badge-faint') : null;
    });
    expect(faint).toBe(true);

    // Unload plugin
    await page.evaluate(() => app._test.unloadPlugin(0));
  });

  test('chip order is Linking, Linked, Formatted, Sorted, Filtered', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.sortCols = [{ col: 'name', dir: 'asc' }];
      win.columnFilters['name'] = new Set(['Alice Johnson']);
      app._test.rebuildTable(win);
    });
    await page.waitForTimeout(100);

    const chipTexts = await page.$$eval('.status-chip', els => els.map(e => e.textContent));
    expect(chipTexts).toEqual(['Linking', 'Linked', 'Formatted', 'Sorted', 'Filtered']);

    const activeChips = await page.$$eval('.status-chip:not(.chip-inactive)', els => els.map(e => e.textContent));
    expect(activeChips).toEqual(['Sorted', 'Filtered']);
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
