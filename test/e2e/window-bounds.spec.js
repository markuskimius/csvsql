const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow } = require('../helpers');

test.describe('Window Boundary Constraints', () => {
  let winEl;

  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('drag cannot move window past left edge', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.el.style.left = '50px';
      win.el.style.top = '50px';
    });

    const titlebar = page.locator('.subwindow .win-titlebar').first();
    const box = await titlebar.boundingBox();
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(0, box.y + box.height / 2, { steps: 10 });
    await page.mouse.up();

    const left = await page.evaluate(() => parseInt(app._test.windows[0].el.style.left));
    expect(left).toBe(0);
  });

  test('drag cannot move window past top edge', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.el.style.left = '50px';
      win.el.style.top = '50px';
    });

    const titlebar = page.locator('.subwindow .win-titlebar').first();
    const box = await titlebar.boundingBox();
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + box.width / 2, 0, { steps: 10 });
    await page.mouse.up();

    const top = await page.evaluate(() => parseInt(app._test.windows[0].el.style.top));
    expect(top).toBe(0);
  });

  test('drag cannot move window past right edge', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.el.style.left = '50px';
      win.el.style.top = '50px';
      win.el.style.width = '400px';
    });

    const titlebar = page.locator('.subwindow .win-titlebar').first();
    const box = await titlebar.boundingBox();
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(9999, box.y + box.height / 2, { steps: 10 });
    await page.mouse.up();

    const { left, width, areaW } = await page.evaluate(() => {
      const win = app._test.windows[0];
      const area = document.getElementById('window-area');
      return {
        left: parseInt(win.el.style.left),
        width: win.el.offsetWidth,
        areaW: area.clientWidth,
      };
    });
    expect(left + width).toBeLessThanOrEqual(areaW);
  });

  test('drag cannot move window past bottom edge', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.el.style.left = '50px';
      win.el.style.top = '50px';
      win.el.style.height = '300px';
    });

    const titlebar = page.locator('.subwindow .win-titlebar').first();
    const box = await titlebar.boundingBox();
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + box.width / 2, 9999, { steps: 10 });
    await page.mouse.up();

    const { top, height, areaH } = await page.evaluate(() => {
      const win = app._test.windows[0];
      const area = document.getElementById('window-area');
      return {
        top: parseInt(win.el.style.top),
        height: win.el.offsetHeight,
        areaH: area.clientHeight,
      };
    });
    expect(top + height).toBeLessThanOrEqual(areaH);
  });

  test('resize right edge cannot exceed workspace', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.el.style.left = '100px';
      win.el.style.top = '50px';
      win.el.style.width = '400px';
      win.el.style.height = '300px';
    });

    const winBox = await page.locator('.subwindow').first().boundingBox();
    const rightEdge = winBox.x + winBox.width;
    await page.mouse.move(rightEdge - 2, winBox.y + winBox.height / 2);
    await page.mouse.down();
    await page.mouse.move(9999, winBox.y + winBox.height / 2, { steps: 10 });
    await page.mouse.up();

    const { left, width, areaW } = await page.evaluate(() => {
      const win = app._test.windows[0];
      const area = document.getElementById('window-area');
      return {
        left: parseInt(win.el.style.left),
        width: parseInt(win.el.style.width),
        areaW: area.clientWidth,
      };
    });
    expect(left + width).toBeLessThanOrEqual(areaW);
  });

  test('resize bottom edge cannot exceed workspace', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.el.style.left = '50px';
      win.el.style.top = '100px';
      win.el.style.width = '400px';
      win.el.style.height = '300px';
    });

    const winBox = await page.locator('.subwindow').first().boundingBox();
    const bottomEdge = winBox.y + winBox.height;
    await page.mouse.move(winBox.x + winBox.width / 2, bottomEdge - 2);
    await page.mouse.down();
    await page.mouse.move(winBox.x + winBox.width / 2, 9999, { steps: 10 });
    await page.mouse.up();

    const { top, height, areaH } = await page.evaluate(() => {
      const win = app._test.windows[0];
      const area = document.getElementById('window-area');
      return {
        top: parseInt(win.el.style.top),
        height: parseInt(win.el.style.height),
        areaH: area.clientHeight,
      };
    });
    expect(top + height).toBeLessThanOrEqual(areaH);
  });

  test('resize left edge cannot go below zero', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.el.style.left = '50px';
      win.el.style.top = '50px';
      win.el.style.width = '400px';
      win.el.style.height = '300px';
    });

    const winBox = await page.locator('.subwindow').first().boundingBox();
    await page.mouse.move(winBox.x + 2, winBox.y + winBox.height / 2);
    await page.mouse.down();
    await page.mouse.move(0, winBox.y + winBox.height / 2, { steps: 10 });
    await page.mouse.up();

    const left = await page.evaluate(() => parseInt(app._test.windows[0].el.style.left));
    expect(left).toBeGreaterThanOrEqual(0);
  });

  test('resize top edge cannot go below zero', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.el.style.left = '50px';
      win.el.style.top = '50px';
      win.el.style.width = '400px';
      win.el.style.height = '300px';
    });

    const winBox = await page.locator('.subwindow').first().boundingBox();
    await page.mouse.move(winBox.x + winBox.width / 2, winBox.y + 2);
    await page.mouse.down();
    await page.mouse.move(winBox.x + winBox.width / 2, 0, { steps: 10 });
    await page.mouse.up();

    const top = await page.evaluate(() => parseInt(app._test.windows[0].el.style.top));
    expect(top).toBeGreaterThanOrEqual(0);
  });

  test('nudge window stops at right boundary', async ({ page }) => {
    const { areaW, winW } = await page.evaluate(() => {
      const win = app._test.windows[0];
      const area = document.getElementById('window-area');
      const areaW = area.clientWidth;
      const winW = win.el.offsetWidth;
      win.el.style.left = (areaW - winW - 2) + 'px';
      win.el.style.top = '50px';
      return { areaW, winW };
    });

    // Focus a cell so Ctrl+L nudge works
    await page.locator('.subwindow .data-cell').first().click();
    await page.keyboard.press('Control+l');
    await page.keyboard.press('Control+l');
    await page.keyboard.press('Control+l');

    const left = await page.evaluate(() => parseInt(app._test.windows[0].el.style.left));
    const width = await page.evaluate(() => app._test.windows[0].el.offsetWidth);
    expect(left + width).toBeLessThanOrEqual(areaW);
  });

  test('nudge window stops at bottom boundary', async ({ page }) => {
    const { areaH, winH } = await page.evaluate(() => {
      const win = app._test.windows[0];
      const area = document.getElementById('window-area');
      const areaH = area.clientHeight;
      const winH = win.el.offsetHeight;
      win.el.style.left = '50px';
      win.el.style.top = (areaH - winH - 2) + 'px';
      return { areaH, winH };
    });

    await page.locator('.subwindow .data-cell').first().click();
    await page.keyboard.press('Control+j');
    await page.keyboard.press('Control+j');
    await page.keyboard.press('Control+j');

    const top = await page.evaluate(() => parseInt(app._test.windows[0].el.style.top));
    const height = await page.evaluate(() => app._test.windows[0].el.offsetHeight);
    expect(top + height).toBeLessThanOrEqual(areaH);
  });

  test('nudge window stops at left boundary', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.el.style.left = '2px';
      win.el.style.top = '50px';
    });

    await page.locator('.subwindow .data-cell').first().click();
    await page.keyboard.press('Control+h');
    await page.keyboard.press('Control+h');
    await page.keyboard.press('Control+h');

    const left = await page.evaluate(() => parseInt(app._test.windows[0].el.style.left));
    expect(left).toBe(0);
  });

  test('nudge window stops at top boundary', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.el.style.left = '50px';
      win.el.style.top = '2px';
    });

    await page.locator('.subwindow .data-cell').first().click();
    await page.keyboard.press('Control+k');
    await page.keyboard.press('Control+k');
    await page.keyboard.press('Control+k');

    const top = await page.evaluate(() => parseInt(app._test.windows[0].el.style.top));
    expect(top).toBe(0);
  });

  test('window stays within bounds after drag to corner', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.el.style.left = '50px';
      win.el.style.top = '50px';
      win.el.style.width = '400px';
      win.el.style.height = '300px';
    });

    const titlebar = page.locator('.subwindow .win-titlebar').first();
    const box = await titlebar.boundingBox();
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(9999, 9999, { steps: 10 });
    await page.mouse.up();

    const { left, top, width, height, areaW, areaH } = await page.evaluate(() => {
      const win = app._test.windows[0];
      const area = document.getElementById('window-area');
      return {
        left: parseInt(win.el.style.left),
        top: parseInt(win.el.style.top),
        width: win.el.offsetWidth,
        height: win.el.offsetHeight,
        areaW: area.clientWidth,
        areaH: area.clientHeight,
      };
    });
    expect(left).toBeGreaterThanOrEqual(0);
    expect(top).toBeGreaterThanOrEqual(0);
    expect(left + width).toBeLessThanOrEqual(areaW);
    expect(top + height).toBeLessThanOrEqual(areaH);
  });
});
