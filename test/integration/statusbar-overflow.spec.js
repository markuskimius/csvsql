const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow } = require('../helpers');

test.describe('Statusbar Overflow', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, 'sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('statusbar has white-space nowrap to prevent text wrapping', async ({ page }) => {
    const whiteSpace = await page.$eval('.win-statusbar', el =>
      getComputedStyle(el).whiteSpace
    );
    expect(whiteSpace).toBe('nowrap');
  });

  test('statusbar has overflow hidden to prevent content leaking', async ({ page }) => {
    const overflow = await page.$eval('.win-statusbar', el =>
      getComputedStyle(el).overflow
    );
    expect(overflow).toBe('hidden');
  });

  test('status-left has overflow hidden and text-overflow ellipsis', async ({ page }) => {
    const styles = await page.$eval('.status-left', el => {
      const cs = getComputedStyle(el);
      return { overflow: cs.overflow, textOverflow: cs.textOverflow };
    });
    expect(styles.overflow).toBe('hidden');
    expect(styles.textOverflow).toBe('ellipsis');
  });

  test('status-right has overflow hidden and text-overflow ellipsis', async ({ page }) => {
    const styles = await page.$eval('.status-right', el => {
      const cs = getComputedStyle(el);
      return { overflow: cs.overflow, textOverflow: cs.textOverflow };
    });
    expect(styles.overflow).toBe('hidden');
    expect(styles.textOverflow).toBe('ellipsis');
  });

  test('status-left and status-right allow shrinking (flex-shrink > 0)', async ({ page }) => {
    const shrinks = await page.evaluate(() => {
      const left = document.querySelector('.status-left');
      const right = document.querySelector('.status-right');
      const cs = (el) => getComputedStyle(el).flexShrink;
      return { left: cs(left), right: cs(right) };
    });
    expect(Number(shrinks.left)).toBeGreaterThan(0);
    expect(Number(shrinks.right)).toBeGreaterThan(0);
  });

  test('status-center does not shrink (flex-shrink 0)', async ({ page }) => {
    const shrink = await page.$eval('.status-center', el =>
      getComputedStyle(el).flexShrink
    );
    expect(shrink).toBe('0');
  });

  test('statusbar does not exceed window width when window is narrow', async ({ page }) => {
    const winEl = await page.$('.subwindow');
    const box = await winEl.boundingBox();

    // Resize the window to be very narrow (120px)
    await page.evaluate(id => {
      const win = app._test.windows.find(w => w.id === id);
      const el = win.el;
      el.style.width = '120px';
      el.style.left = '10px';
    }, await page.$eval('.subwindow', el => parseInt(el.dataset.winId || el.id.replace('win-', ''))));

    // Wait for layout to settle
    await page.waitForTimeout(100);

    const result = await page.evaluate(() => {
      const statusbar = document.querySelector('.win-statusbar');
      const win = statusbar.closest('.subwindow');
      const winRect = win.getBoundingClientRect();
      const statusRect = statusbar.getBoundingClientRect();
      return {
        statusRight: statusRect.right,
        winRight: winRect.right,
        statusLeft: statusRect.left,
        winLeft: winRect.left,
      };
    });

    expect(result.statusRight).toBeLessThanOrEqual(result.winRight + 1);
    expect(result.statusLeft).toBeGreaterThanOrEqual(result.winLeft - 1);
  });

  test('statusbar stays single-line when window is narrow', async ({ page }) => {
    // Resize the window to be narrow
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.el.style.width = '150px';
    });
    await page.waitForTimeout(100);

    const statusbarHeight = await page.$eval('.win-statusbar', el =>
      el.getBoundingClientRect().height
    );

    // A single-line statusbar at 11px font should be well under 30px
    expect(statusbarHeight).toBeLessThan(30);
  });

  test('statusbar text content is present for row and column counts', async ({ page }) => {
    const texts = await page.evaluate(() => {
      const left = document.querySelector('.status-left');
      const right = document.querySelector('.status-right');
      return { left: left.textContent, right: right.textContent };
    });

    expect(texts.left).toMatch(/\d+ of \d+ rows/);
    expect(texts.right).toMatch(/\d+ columns/);
  });

  test('status-left and status-right have min-width 0 for flex truncation', async ({ page }) => {
    const minWidths = await page.evaluate(() => {
      const left = document.querySelector('.status-left');
      const right = document.querySelector('.status-right');
      const cs = (el) => getComputedStyle(el).minWidth;
      return { left: cs(left), right: cs(right) };
    });
    expect(minWidths.left).toBe('0px');
    expect(minWidths.right).toBe('0px');
  });
});
