const { test, expect } = require('@playwright/test');
const { openApp } = require('../helpers');

// Scrollbars are styled once globally: 12px gutter, 6px visible thumb
// centered by a 3px transparent border. The 3px transparent margin is what
// keeps the window resize handles' 3px inward overhang off the visible
// thumb (see CLAUDE.md "Scrollbars"). Headless Chromium uses overlay
// scrollbars, so these tests assert the stylesheet rules and the handle
// geometry rather than rendered gutter widths.

// Collect every style rule (flattening media rules) whose selector mentions
// ::-webkit-scrollbar, as { selector, media, style: {prop: value} }.
async function scrollbarRules(page) {
  return page.evaluate(() => {
    const out = [];
    const visit = (rules, media) => {
      for (const rule of rules) {
        if (rule.media) { visit(rule.cssRules, rule.media.mediaText); continue; }
        if (!rule.selectorText || !rule.selectorText.includes('::-webkit-scrollbar')) continue;
        const style = {};
        for (const prop of rule.style) style[prop] = rule.style.getPropertyValue(prop);
        out.push({ selector: rule.selectorText, media: media || null, style });
      }
    };
    for (const sheet of document.styleSheets) visit(sheet.cssRules, null);
    return out;
  });
}

test.describe('Uniform scrollbars', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
  });

  test('one global scrollbar rule: 12px gutter, thumb border matches handle overhang', async ({ page }) => {
    const rules = await scrollbarRules(page);

    const size = rules.find(r => r.selector === '::-webkit-scrollbar' && !r.media);
    expect(size, 'global ::-webkit-scrollbar rule missing').toBeTruthy();
    expect(size.style['width']).toBe('12px');
    expect(size.style['height']).toBe('12px');

    const thumb = rules.find(r => r.selector === '::-webkit-scrollbar-thumb' && !r.media);
    expect(thumb, 'global ::-webkit-scrollbar-thumb rule missing').toBeTruthy();
    // 3px transparent border + padding-box clip = the collision-free margin
    expect(thumb.style['border-top-width']).toBe('3px');
    expect(thumb.style['border-top-color']).toBe('transparent');
    expect(thumb.style['background-clip']).toBe('padding-box');
  });

  test('element-scoped scrollbar rules only hide scrollbars', async ({ page }) => {
    const rules = await scrollbarRules(page);
    for (const rule of rules) {
      for (const sel of rule.selector.split(',').map(s => s.trim())) {
        if (sel.startsWith('::-webkit-scrollbar')) continue; // global
        expect(rule.style['display'],
          `scoped scrollbar rule "${sel}" must only hide (display: none) — ` +
          'style scrollbars globally instead (CLAUDE.md "Scrollbars")'
        ).toBe('none');
      }
    }
  });

  test('coarse pointers get a 16px gutter with a wider thumb border', async ({ page }) => {
    const rules = await scrollbarRules(page);
    const coarse = rules.filter(r => r.media && r.media.includes('pointer: coarse'));
    const size = coarse.find(r => r.selector === '::-webkit-scrollbar');
    expect(size, 'coarse-pointer ::-webkit-scrollbar override missing').toBeTruthy();
    expect(size.style['width']).toBe('16px');
    expect(size.style['height']).toBe('16px');

    const thumb = coarse.find(r => r.selector.includes('::-webkit-scrollbar-thumb'));
    expect(thumb, 'coarse-pointer thumb override missing').toBeTruthy();
    expect(thumb.style['border-top-width']).toBe('4px');
  });

  test('resize handles overhang exactly the thumb border into the window', async ({ page }) => {
    await page.evaluate(() => app.loadExampleData());
    await page.waitForSelector('.subwindow .table-container');

    const intrusion = await page.evaluate(() => {
      const sw = [...document.querySelectorAll('.subwindow')]
        .sort((a, b) => (+b.style.zIndex || 0) - (+a.style.zIndex || 0))[0];
      const tc = sw.querySelector('.table-container').getBoundingClientRect();
      const rh = sw.querySelector('.resize-handle.rh-right').getBoundingClientRect();
      return tc.right - rh.left;
    });
    // Must not exceed the thumb's 3px transparent border, or the resize
    // cursor covers the visible scrollbar thumb — and must stay >0 so a
    // window flush against the workspace edge is still resizable from inside.
    expect(intrusion).toBeLessThanOrEqual(3);
    expect(intrusion).toBeGreaterThan(0);
  });
});
