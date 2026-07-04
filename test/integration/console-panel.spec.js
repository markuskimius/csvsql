const { test, expect } = require('@playwright/test');
const { openApp } = require('../helpers');

test.describe('Console Panel', () => {
  test('default height is 113px and fits exactly 3 lines of SQL', async ({ page }) => {
    await openApp(page);

    const m = await page.evaluate(() => {
      document.getElementById('console-status').textContent = 'Ready';
      const input = document.getElementById('sql-input');
      input.value = 'line1\nline2\nline3';
      const cs = getComputedStyle(input);
      const padding = parseFloat(cs.paddingTop) + parseFloat(cs.paddingBottom);
      return {
        panelHeight: document.getElementById('console-panel').getBoundingClientRect().height,
        textHeight: input.getBoundingClientRect().height - padding,
        scrolls: input.scrollHeight > input.clientHeight,
      };
    });

    expect(m.panelHeight).toBe(113);
    // 3 lines at 15px line height, no scrollbar
    expect(m.textHeight).toBe(45);
    expect(m.scrolls).toBe(false);
  });

  test('4th line overflows the default height', async ({ page }) => {
    await openApp(page);

    const scrolls = await page.evaluate(() => {
      const input = document.getElementById('sql-input');
      input.value = 'line1\nline2\nline3\nline4';
      return input.scrollHeight > input.clientHeight;
    });

    expect(scrolls).toBe(true);
  });

  test('drag on resize handle grows the panel', async ({ page }) => {
    await openApp(page);

    const handle = page.locator('#console-resize-handle');
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2 - 100);
    await page.mouse.up();

    const height = await page.evaluate(() =>
      document.getElementById('console-panel').offsetHeight);
    expect(height).toBeGreaterThan(113);
    expect(height).toBeLessThanOrEqual(113 + 110);
  });

  test('resize clamps to 60px minimum', async ({ page }) => {
    await openApp(page);

    const handle = page.locator('#console-resize-handle');
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + box.width / 2, box.y + 2000);
    await page.mouse.up();

    const height = await page.evaluate(() =>
      document.getElementById('console-panel').offsetHeight);
    expect(height).toBe(60);
  });

  test('Run, History, and Clear buttons are equal width and labels do not wrap', async ({ page }) => {
    await openApp(page);
    const buttons = page.locator('#console-actions button:not(#btn-interrupt)');
    await expect(buttons).toHaveCount(3);
    await expect(buttons.nth(0)).toHaveText('Run');
    await expect(buttons.nth(1)).toHaveText('History');
    await expect(buttons.nth(2)).toHaveText('Clear');

    const boxes = await buttons.evaluateAll(els => els.map(el => {
      const r = el.getBoundingClientRect();
      return { width: r.width, height: r.height, whiteSpace: getComputedStyle(el).whiteSpace };
    }));
    for (const box of boxes) {
      expect(box.width).toBeCloseTo(boxes[0].width, 1);
      // Single-line labels: equal heights, none tall enough to hold two lines
      expect(box.height).toBeCloseTo(boxes[0].height, 1);
      expect(box.height).toBeLessThan(30);
      expect(box.whiteSpace).toBe('nowrap');
    }
  });

  test('Run stays the first console action button', async ({ page }) => {
    // The executeSQL test helper clicks the first button in #console-actions
    await openApp(page);
    await expect(page.locator('#console-actions button:first-child')).toHaveText('Run');
  });

  test('buttons stay equal width and unwrapped when Interrupt appears', async ({ page }) => {
    await openApp(page);
    // Mirror showInterruptButton(): it appends #btn-interrupt to #console-actions
    await page.evaluate(() => {
      const btn = document.createElement('button');
      btn.id = 'btn-interrupt';
      btn.textContent = 'Interrupt';
      document.getElementById('console-actions').appendChild(btn);
    });

    const boxes = await page.locator('#console-actions button').evaluateAll(els => els.map(el => {
      const r = el.getBoundingClientRect();
      return { width: r.width, height: r.height, whiteSpace: getComputedStyle(el).whiteSpace };
    }));
    expect(boxes.length).toBe(4);
    for (const box of boxes) {
      expect(box.width).toBeCloseTo(boxes[0].width, 1);
      expect(box.height).toBeLessThan(30);
      expect(box.whiteSpace).toBe('nowrap');
    }
  });

  test('resize clamps to 50% of viewport maximum', async ({ page }) => {
    await openApp(page);

    const handle = page.locator('#console-resize-handle');
    const box = await handle.boundingBox();

    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + box.width / 2, 0);
    await page.mouse.up();

    const m = await page.evaluate(() => ({
      height: document.getElementById('console-panel').offsetHeight,
      maxAllowed: Math.round(window.innerHeight * 0.5),
    }));
    expect(m.height).toBeLessThanOrEqual(m.maxAllowed + 1);
    expect(m.height).toBeGreaterThan(113);
  });
});
