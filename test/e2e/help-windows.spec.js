const { test, expect } = require('@playwright/test');
const { openApp } = require('../helpers');

test.describe('Help Windows', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
  });

  test('manual pre elements have correct height for content', async ({ page }) => {
    await page.evaluate(() => app.showManual());
    await page.waitForSelector('.help-body');

    const preInfo = await page.evaluate(() => {
      const pres = document.querySelectorAll('.help-body pre');
      return [...pres].map(pre => {
        const lines = pre.textContent.trim().split('\n').length;
        const height = pre.getBoundingClientRect().height;
        const style = getComputedStyle(pre);
        const fontSize = parseFloat(style.fontSize);
        const lineHeight = parseFloat(style.lineHeight) || fontSize * 1.6;
        const paddingTop = parseFloat(style.paddingTop);
        const paddingBottom = parseFloat(style.paddingBottom);
        const minExpected = lines * fontSize + paddingTop + paddingBottom;
        return { lines, height, minExpected, fontSize };
      });
    });

    expect(preInfo.length).toBeGreaterThan(0);
    for (const info of preInfo) {
      expect(info.height).toBeGreaterThanOrEqual(info.minExpected);
    }
  });

  test('manual pre elements are not collapsed by flex layout', async ({ page }) => {
    await page.evaluate(() => app.showManual());
    await page.waitForSelector('.help-body');

    const allValid = await page.evaluate(() => {
      const pres = document.querySelectorAll('.help-body pre');
      return [...pres].every(pre => {
        const height = pre.getBoundingClientRect().height;
        const fontSize = parseFloat(getComputedStyle(pre).fontSize);
        return height >= fontSize;
      });
    });

    expect(allValid).toBe(true);
  });

  test('multi-line pre elements are taller than single-line ones', async ({ page }) => {
    await page.evaluate(() => app.showManual());
    await page.waitForSelector('.help-body');

    const heights = await page.evaluate(() => {
      const pres = document.querySelectorAll('.help-body pre');
      return [...pres].map(pre => ({
        lines: pre.textContent.trim().split('\n').length,
        height: pre.getBoundingClientRect().height,
      }));
    });

    const multiLine = heights.filter(h => h.lines > 1);
    const singleLine = heights.filter(h => h.lines === 1);
    if (multiLine.length > 0 && singleLine.length > 0) {
      const avgMulti = multiLine.reduce((s, h) => s + h.height, 0) / multiLine.length;
      const avgSingle = singleLine.reduce((s, h) => s + h.height, 0) / singleLine.length;
      expect(avgMulti).toBeGreaterThan(avgSingle);
    }
  });

  test('manual is a singleton window', async ({ page }) => {
    await page.evaluate(() => app.showManual());
    await page.waitForSelector('.help-body');
    await page.evaluate(() => app.showManual());
    await page.waitForTimeout(200);

    const manualCount = await page.evaluate(() => {
      return [...document.querySelectorAll('.win-title')]
        .filter(t => t.textContent.includes("User's Manual")).length;
    });
    expect(manualCount).toBe(1);
  });

  test('about window is a singleton', async ({ page }) => {
    await page.evaluate(() => app.showAbout());
    await page.waitForSelector('.help-body');
    await page.evaluate(() => app.showAbout());
    await page.waitForTimeout(200);

    const aboutCount = await page.evaluate(() => {
      return [...document.querySelectorAll('.win-title')]
        .filter(t => t.textContent.includes('About')).length;
    });
    expect(aboutCount).toBe(1);
  });

  test('help windows are not dockable', async ({ page }) => {
    await page.evaluate(() => app.showManual());
    await page.waitForSelector('.help-body');

    const dockable = await page.evaluate(() => {
      const win = app._test.windows.find(w => w.title === "User's Manual");
      return win ? win.dockable : null;
    });
    expect(dockable).toBe(false);
  });

  test('pre elements have flex-shrink 0', async ({ page }) => {
    await page.evaluate(() => app.showManual());
    await page.waitForSelector('.help-body pre');

    const flexShrink = await page.evaluate(() => {
      const pre = document.querySelector('.help-body pre');
      return getComputedStyle(pre).flexShrink;
    });
    expect(flexShrink).toBe('0');
  });

  test('help-body is a flex column container', async ({ page }) => {
    await page.evaluate(() => app.showManual());
    await page.waitForSelector('.help-body');

    const styles = await page.evaluate(() => {
      const el = document.querySelector('.help-body');
      const s = getComputedStyle(el);
      return { display: s.display, flexDirection: s.flexDirection };
    });
    expect(styles.display).toBe('flex');
    expect(styles.flexDirection).toBe('column');
  });

  test('manual contains expected sections', async ({ page }) => {
    await page.evaluate(() => app.showManual());
    await page.waitForSelector('.help-body');

    const headings = await page.evaluate(() => {
      return [...document.querySelectorAll('.help-body h4')].map(h => h.textContent);
    });
    expect(headings).toContain('SQL Syntax Reference');
    expect(headings).toContain('REGEXP');
    expect(headings).toContain('SELECT INTO');
  });

  test('pre elements scroll horizontally for long content', async ({ page }) => {
    await page.evaluate(() => app.showManual());
    await page.waitForSelector('.help-body pre');

    const overflowX = await page.evaluate(() => {
      const pre = document.querySelector('.help-body pre');
      return getComputedStyle(pre).overflowX;
    });
    expect(overflowX).toBe('auto');
  });
});
