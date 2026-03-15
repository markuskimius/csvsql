const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, getWindowCount } = require('../helpers');

test.describe('Window Management', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    // Open two files to have multiple windows
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
    await uploadFile(page, '../test/sample2.psv');
    await waitForWindow(page, 'sample2');
  });

  test('opening files creates windows', async ({ page }) => {
    const count = await getWindowCount(page);
    expect(count).toBe(2);
  });

  test('close removes window', async ({ page }) => {
    // Set up dialog handler BEFORE clicking close
    page.on('dialog', dialog => dialog.accept());
    const closeBtn = page.locator('.subwindow .btn-close').first();
    await closeBtn.click();
    await page.waitForTimeout(500);
    const count = await getWindowCount(page);
    expect(count).toBe(1);
  });

  test('tile horizontally arranges windows', async ({ page }) => {
    await page.evaluate(() => app.layoutTileH());
    await page.waitForTimeout(300);

    const positions = await page.evaluate(() => {
      return [...document.querySelectorAll('.subwindow:not(.minimized)')].map(el => ({
        left: parseInt(el.style.left),
        top: parseInt(el.style.top),
        width: parseInt(el.style.width),
        height: parseInt(el.style.height),
      }));
    });

    expect(positions.length).toBe(2);
    // Horizontal tile: same width, stacked vertically
    positions.forEach(p => {
      expect(p.width).toBeGreaterThan(0);
      expect(p.height).toBeGreaterThan(0);
    });
  });

  test('tile vertically arranges windows', async ({ page }) => {
    await page.evaluate(() => app.layoutTileV());
    await page.waitForTimeout(300);

    const positions = await page.evaluate(() => {
      return [...document.querySelectorAll('.subwindow:not(.minimized)')].map(el => ({
        left: parseInt(el.style.left),
        width: parseInt(el.style.width),
      }));
    });

    expect(positions.length).toBe(2);
    // Vertical tile: side by side, different left positions
    expect(positions[0].left).toBe(0);
    expect(positions[1].left).toBeGreaterThan(0);
  });

  test('grid arranges windows', async ({ page }) => {
    await page.evaluate(() => app.layoutGrid());
    await page.waitForTimeout(300);

    const positions = await page.evaluate(() => {
      return [...document.querySelectorAll('.subwindow:not(.minimized)')].map(el => ({
        width: parseInt(el.style.width),
        height: parseInt(el.style.height),
      }));
    });

    expect(positions.length).toBe(2);
    positions.forEach(p => {
      expect(p.width).toBeGreaterThan(0);
      expect(p.height).toBeGreaterThan(0);
    });
  });

  test('cascade offsets windows', async ({ page }) => {
    await page.evaluate(() => app.layoutCascade());
    await page.waitForTimeout(300);

    const positions = await page.evaluate(() => {
      return [...document.querySelectorAll('.subwindow:not(.minimized)')].map(el => ({
        left: parseInt(el.style.left),
        top: parseInt(el.style.top),
      }));
    });

    expect(positions.length).toBe(2);
    // Cascade should offset each window by 30px
    expect(positions[1].left).toBeGreaterThan(positions[0].left);
    expect(positions[1].top).toBeGreaterThan(positions[0].top);
  });

  test('minimize all hides windows', async ({ page }) => {
    await page.evaluate(() => app.minimizeAll());
    await page.waitForTimeout(300);

    const minimizedCount = await page.evaluate(() => {
      return document.querySelectorAll('.subwindow.minimized').length;
    });
    expect(minimizedCount).toBe(2);
  });

  test('restore all shows minimized windows', async ({ page }) => {
    await page.evaluate(() => app.minimizeAll());
    await page.waitForTimeout(200);
    await page.evaluate(() => app.restoreAll());
    await page.waitForTimeout(200);

    const minimizedCount = await page.evaluate(() => {
      return document.querySelectorAll('.subwindow.minimized').length;
    });
    expect(minimizedCount).toBe(0);
  });
});
