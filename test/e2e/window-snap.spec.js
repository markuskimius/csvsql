const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow } = require('../helpers');

test.describe('Window Snap — Drag', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('drag snaps left edge to area left', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.el.style.left = '100px';
      win.el.style.top = '100px';
      win.el.style.width = '400px';
      win.el.style.height = '300px';
    });

    const titlebar = page.locator('.subwindow .win-titlebar').first();
    const box = await titlebar.boundingBox();
    const startX = box.x + box.width / 2;
    const startY = box.y + box.height / 2;
    await page.mouse.move(startX, startY);
    await page.mouse.down();
    // Move left by 93px: 100 - 93 = 7, within SNAP_THRESHOLD (10) of 0
    await page.mouse.move(startX - 93, startY, { steps: 10 });
    await page.mouse.up();

    const left = await page.evaluate(() => parseInt(app._test.windows[0].el.style.left));
    expect(left).toBe(0);
  });

  test('drag snaps top edge to area top', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.el.style.left = '100px';
      win.el.style.top = '100px';
      win.el.style.width = '400px';
      win.el.style.height = '300px';
    });

    const titlebar = page.locator('.subwindow .win-titlebar').first();
    const box = await titlebar.boundingBox();
    const startX = box.x + box.width / 2;
    const startY = box.y + box.height / 2;
    await page.mouse.move(startX, startY);
    await page.mouse.down();
    // Move up by 93px: 100 - 93 = 7, within SNAP_THRESHOLD (10) of 0
    await page.mouse.move(startX, startY - 93, { steps: 10 });
    await page.mouse.up();

    const top = await page.evaluate(() => parseInt(app._test.windows[0].el.style.top));
    expect(top).toBe(0);
  });

  test('drag snaps right edge to area right', async ({ page }) => {
    const setup = await page.evaluate(() => {
      const win = app._test.windows[0];
      const area = document.getElementById('window-area');
      const areaW = area.clientWidth;
      // Position so dragging right by ~50px puts the right edge within snap range
      win.el.style.width = '400px';
      win.el.style.height = '300px';
      win.el.style.left = (areaW - 400 - 50) + 'px';
      win.el.style.top = '100px';
      return { areaW };
    });

    const titlebar = page.locator('.subwindow .win-titlebar').first();
    const box = await titlebar.boundingBox();
    const startX = box.x + box.width / 2;
    const startY = box.y + box.height / 2;
    await page.mouse.move(startX, startY);
    await page.mouse.down();
    // Move right by 45px: right edge would be areaW - 5, within SNAP_THRESHOLD of areaW
    await page.mouse.move(startX + 45, startY, { steps: 10 });
    await page.mouse.up();

    const { left, width } = await page.evaluate(() => {
      const win = app._test.windows[0];
      return { left: parseInt(win.el.style.left), width: win.el.offsetWidth };
    });
    expect(left + width).toBe(setup.areaW);
  });

  test('drag snaps bottom edge to area bottom', async ({ page }) => {
    const setup = await page.evaluate(() => {
      const win = app._test.windows[0];
      const area = document.getElementById('window-area');
      const areaH = area.clientHeight;
      win.el.style.width = '400px';
      win.el.style.height = '200px';
      win.el.style.left = '100px';
      win.el.style.top = (areaH - 200 - 50) + 'px';
      return { areaH };
    });

    const titlebar = page.locator('.subwindow .win-titlebar').first();
    const box = await titlebar.boundingBox();
    const startX = box.x + box.width / 2;
    const startY = box.y + box.height / 2;
    await page.mouse.move(startX, startY);
    await page.mouse.down();
    await page.mouse.move(startX, startY + 45, { steps: 10 });
    await page.mouse.up();

    const { top, height } = await page.evaluate(() => {
      const win = app._test.windows[0];
      return { top: parseInt(win.el.style.top), height: win.el.offsetHeight };
    });
    expect(top + height).toBe(setup.areaH);
  });

  test('no snap when outside threshold', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.el.style.left = '200px';
      win.el.style.top = '200px';
      win.el.style.width = '300px';
      win.el.style.height = '200px';
    });

    const titlebar = page.locator('.subwindow .win-titlebar').first();
    const box = await titlebar.boundingBox();
    const startX = box.x + box.width / 2;
    const startY = box.y + box.height / 2;
    await page.mouse.move(startX, startY);
    await page.mouse.down();
    // Move to left=100, far from any snap edge (only edges are 0 and areaW/areaH)
    await page.mouse.move(startX - 100, startY - 100, { steps: 10 });
    await page.mouse.up();

    const { left, top } = await page.evaluate(() => {
      const win = app._test.windows[0];
      return { left: parseInt(win.el.style.left), top: parseInt(win.el.style.top) };
    });
    expect(left).toBe(100);
    expect(top).toBe(100);
  });
});

test.describe('Window Snap — Multi-window drag', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
    await uploadFile(page, '../test/sample2.psv');
    await waitForWindow(page, 'sample2');
  });

  test('drag snaps left edge to other window right edge', async ({ page }) => {
    await page.evaluate(() => {
      const [w0, w1] = app._test.windows;
      // Window 0 (sample1): fixed at left=0, width=400 → right edge at 400
      w0.el.style.left = '0px';
      w0.el.style.top = '0px';
      w0.el.style.width = '400px';
      w0.el.style.height = '300px';
      // Window 1 (sample2): start at left=500
      w1.el.style.left = '500px';
      w1.el.style.top = '0px';
      w1.el.style.width = '300px';
      w1.el.style.height = '300px';
    });

    // Drag window 1 left so its left edge lands near 400 (window 0's right edge)
    const titlebar = page.locator('.subwindow .win-titlebar').nth(1);
    const box = await titlebar.boundingBox();
    const startX = box.x + box.width / 2;
    const startY = box.y + box.height / 2;
    await page.mouse.move(startX, startY);
    await page.mouse.down();
    // Move left by 95px: 500 - 95 = 405, within 10px of 400
    await page.mouse.move(startX - 95, startY, { steps: 10 });
    await page.mouse.up();

    const left = await page.evaluate(() => parseInt(app._test.windows[1].el.style.left));
    expect(left).toBe(400);
  });

  test('drag snaps right edge to other window left edge', async ({ page }) => {
    await page.evaluate(() => {
      const [w0, w1] = app._test.windows;
      // Window 0: fixed at left=500
      w0.el.style.left = '500px';
      w0.el.style.top = '0px';
      w0.el.style.width = '300px';
      w0.el.style.height = '300px';
      // Window 1: start at left=0, width=300 → right edge at 300
      w1.el.style.left = '0px';
      w1.el.style.top = '0px';
      w1.el.style.width = '300px';
      w1.el.style.height = '300px';
    });

    // Drag window 1 right so its right edge (left+300) lands near 500 (window 0's left)
    const titlebar = page.locator('.subwindow .win-titlebar').nth(1);
    const box = await titlebar.boundingBox();
    const startX = box.x + box.width / 2;
    const startY = box.y + box.height / 2;
    await page.mouse.move(startX, startY);
    await page.mouse.down();
    // Move right by 195px: left=195, right=495, within 10px of 500
    await page.mouse.move(startX + 195, startY, { steps: 10 });
    await page.mouse.up();

    const { left, width } = await page.evaluate(() => {
      const win = app._test.windows[1];
      return { left: parseInt(win.el.style.left), width: win.el.offsetWidth };
    });
    expect(left + width).toBe(500);
  });

  test('drag snaps top edge to other window bottom edge', async ({ page }) => {
    await page.evaluate(() => {
      const [w0, w1] = app._test.windows;
      // Window 0: fixed at top=0, height=200 → bottom edge at 200
      w0.el.style.left = '0px';
      w0.el.style.top = '0px';
      w0.el.style.width = '400px';
      w0.el.style.height = '200px';
      // Window 1: start at top=300
      w1.el.style.left = '0px';
      w1.el.style.top = '300px';
      w1.el.style.width = '400px';
      w1.el.style.height = '200px';
    });

    const titlebar = page.locator('.subwindow .win-titlebar').nth(1);
    const box = await titlebar.boundingBox();
    const startX = box.x + box.width / 2;
    const startY = box.y + box.height / 2;
    await page.mouse.move(startX, startY);
    await page.mouse.down();
    // Move up by 95px: 300 - 95 = 205, within 10px of 200
    await page.mouse.move(startX, startY - 95, { steps: 10 });
    await page.mouse.up();

    const top = await page.evaluate(() => parseInt(app._test.windows[1].el.style.top));
    expect(top).toBe(200);
  });

  test('drag aligns left edges of two windows', async ({ page }) => {
    await page.evaluate(() => {
      const [w0, w1] = app._test.windows;
      // Window 0: fixed at left=200
      w0.el.style.left = '200px';
      w0.el.style.top = '0px';
      w0.el.style.width = '300px';
      w0.el.style.height = '200px';
      // Window 1: start at left=100
      w1.el.style.left = '100px';
      w1.el.style.top = '220px';
      w1.el.style.width = '300px';
      w1.el.style.height = '200px';
    });

    const titlebar = page.locator('.subwindow .win-titlebar').nth(1);
    const box = await titlebar.boundingBox();
    const startX = box.x + box.width / 2;
    const startY = box.y + box.height / 2;
    await page.mouse.move(startX, startY);
    await page.mouse.down();
    // Move right by 95px: 100 + 95 = 195, within 10px of 200
    await page.mouse.move(startX + 95, startY, { steps: 10 });
    await page.mouse.up();

    const left = await page.evaluate(() => parseInt(app._test.windows[1].el.style.left));
    expect(left).toBe(200);
  });
});

test.describe('Window Snap — Resize', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('resize right edge snaps to area right', async ({ page }) => {
    const setup = await page.evaluate(() => {
      const win = app._test.windows[0];
      const area = document.getElementById('window-area');
      const areaW = area.clientWidth;
      win.el.style.left = '100px';
      win.el.style.top = '50px';
      win.el.style.width = (areaW - 100 - 50) + 'px';
      win.el.style.height = '300px';
      return { areaW };
    });

    const winBox = await page.locator('.subwindow').first().boundingBox();
    const rightEdge = winBox.x + winBox.width;
    await page.mouse.move(rightEdge - 2, winBox.y + winBox.height / 2);
    await page.mouse.down();
    // Drag right by 45px: right edge now at areaW - 5, within snap of areaW
    await page.mouse.move(rightEdge + 43, winBox.y + winBox.height / 2, { steps: 10 });
    await page.mouse.up();

    const { left, width } = await page.evaluate(() => {
      const win = app._test.windows[0];
      return { left: parseInt(win.el.style.left), width: parseInt(win.el.style.width) };
    });
    expect(left + width).toBe(setup.areaW);
  });

  test('resize bottom edge snaps to area bottom', async ({ page }) => {
    const setup = await page.evaluate(() => {
      const win = app._test.windows[0];
      const area = document.getElementById('window-area');
      const areaH = area.clientHeight;
      win.el.style.left = '50px';
      win.el.style.top = '50px';
      win.el.style.width = '400px';
      win.el.style.height = (areaH - 50 - 50) + 'px';
      return { areaH };
    });

    const winBox = await page.locator('.subwindow').first().boundingBox();
    const bottomEdge = winBox.y + winBox.height;
    await page.mouse.move(winBox.x + winBox.width / 2, bottomEdge - 2);
    await page.mouse.down();
    await page.mouse.move(winBox.x + winBox.width / 2, bottomEdge + 43, { steps: 10 });
    await page.mouse.up();

    const { top, height } = await page.evaluate(() => {
      const win = app._test.windows[0];
      return { top: parseInt(win.el.style.top), height: parseInt(win.el.style.height) };
    });
    expect(top + height).toBe(setup.areaH);
  });

  test('resize left edge snaps to area left', async ({ page }) => {
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
    // Drag left to x near area left: left would be ~5, within snap of 0
    await page.mouse.move(winBox.x - 43, winBox.y + winBox.height / 2, { steps: 10 });
    await page.mouse.up();

    const left = await page.evaluate(() => parseInt(app._test.windows[0].el.style.left));
    expect(left).toBe(0);
  });

  test('resize top edge snaps to area top', async ({ page }) => {
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
    await page.mouse.move(winBox.x + winBox.width / 2, winBox.y - 43, { steps: 10 });
    await page.mouse.up();

    const top = await page.evaluate(() => parseInt(app._test.windows[0].el.style.top));
    expect(top).toBe(0);
  });
});

test.describe('Window Snap — Resize to other window', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
    await uploadFile(page, '../test/sample2.psv');
    await waitForWindow(page, 'sample2');
  });

  test('resize right edge snaps to other window left edge', async ({ page }) => {
    await page.evaluate(() => {
      const [w0, w1] = app._test.windows;
      // Window 0: will be resized
      w0.el.style.left = '0px';
      w0.el.style.top = '50px';
      w0.el.style.width = '300px';
      w0.el.style.height = '300px';
      // Window 1: target edge at left=500
      w1.el.style.left = '500px';
      w1.el.style.top = '50px';
      w1.el.style.width = '300px';
      w1.el.style.height = '300px';
    });

    const winBox = await page.locator('.subwindow').first().boundingBox();
    const rightEdge = winBox.x + winBox.width;
    await page.mouse.move(rightEdge - 2, winBox.y + winBox.height / 2);
    await page.mouse.down();
    // Resize right to 495px (within snap of 500)
    await page.mouse.move(rightEdge + 192, winBox.y + winBox.height / 2, { steps: 10 });
    await page.mouse.up();

    const width = await page.evaluate(() => parseInt(app._test.windows[0].el.style.width));
    expect(width).toBe(500);
  });
});

test.describe('Window Snap — Guides', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('snap guide lines appear during drag near edge', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.el.style.left = '100px';
      win.el.style.top = '100px';
      win.el.style.width = '400px';
      win.el.style.height = '300px';
    });

    const titlebar = page.locator('.subwindow .win-titlebar').first();
    const box = await titlebar.boundingBox();
    const startX = box.x + box.width / 2;
    const startY = box.y + box.height / 2;
    await page.mouse.move(startX, startY);
    await page.mouse.down();
    // Move within snap threshold of top-left corner (0,0)
    await page.mouse.move(startX - 95, startY - 95, { steps: 10 });

    // Guides should be visible while dragging
    const guideCount = await page.evaluate(() => document.querySelectorAll('.snap-guide').length);
    expect(guideCount).toBeGreaterThan(0);

    await page.mouse.up();

    // Guides should be gone after mouseup
    const guideCountAfter = await page.evaluate(() => document.querySelectorAll('.snap-guide').length);
    expect(guideCountAfter).toBe(0);
  });

  test('no snap guides when not near any edge', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.el.style.left = '200px';
      win.el.style.top = '200px';
      win.el.style.width = '300px';
      win.el.style.height = '200px';
    });

    const titlebar = page.locator('.subwindow .win-titlebar').first();
    const box = await titlebar.boundingBox();
    const startX = box.x + box.width / 2;
    const startY = box.y + box.height / 2;
    await page.mouse.move(startX, startY);
    await page.mouse.down();
    // Small move that stays far from any edge
    await page.mouse.move(startX - 20, startY - 20, { steps: 5 });

    const guideCount = await page.evaluate(() => document.querySelectorAll('.snap-guide').length);
    expect(guideCount).toBe(0);

    await page.mouse.up();
  });

  test('snap guide line appears at correct position', async ({ page }) => {
    await page.evaluate(() => {
      const win = app._test.windows[0];
      win.el.style.left = '100px';
      win.el.style.top = '200px';
      win.el.style.width = '400px';
      win.el.style.height = '300px';
    });

    const titlebar = page.locator('.subwindow .win-titlebar').first();
    const box = await titlebar.boundingBox();
    const startX = box.x + box.width / 2;
    const startY = box.y + box.height / 2;
    await page.mouse.move(startX, startY);
    await page.mouse.down();
    // Drag to snap left edge to 0
    await page.mouse.move(startX - 95, startY, { steps: 10 });

    const guideLeft = await page.evaluate(() => {
      const guide = document.querySelector('.snap-guide-x');
      return guide ? guide.style.left : null;
    });
    expect(guideLeft).toBe('0px');

    await page.mouse.up();
  });
});

test.describe('Window Snap — Unit tests', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('SNAP_THRESHOLD is 10', async ({ page }) => {
    const threshold = await page.evaluate(() => app._test.SNAP_THRESHOLD);
    expect(threshold).toBe(10);
  });

  test('findSnap returns nearest edge within threshold', async ({ page }) => {
    const result = await page.evaluate(() => {
      const { findSnap, SNAP_THRESHOLD } = app._test;
      return {
        snapsTo0: findSnap(5, [0, 100, 200], SNAP_THRESHOLD),
        snapsTo100: findSnap(95, [0, 100, 200], SNAP_THRESHOLD),
        noSnap: findSnap(50, [0, 100, 200], SNAP_THRESHOLD),
        exactMatch: findSnap(100, [0, 100, 200], SNAP_THRESHOLD),
      };
    });
    expect(result.snapsTo0).toBe(0);
    expect(result.snapsTo100).toBe(100);
    expect(result.noSnap).toBeNull();
    expect(result.exactMatch).toBe(100);
  });

  test('findSnap returns closest when multiple edges within threshold', async ({ page }) => {
    const result = await page.evaluate(() => {
      const { findSnap, SNAP_THRESHOLD } = app._test;
      // Value 97 is 3 away from 100, 7 away from 90 — should pick 100
      return findSnap(97, [90, 100], SNAP_THRESHOLD);
    });
    expect(result).toBe(100);
  });

  test('getSnapEdges includes area boundaries and window edges', async ({ page }) => {
    const edges = await page.evaluate(() => {
      const win = app._test.windows[0];
      win.el.style.left = '50px';
      win.el.style.top = '60px';
      win.el.style.width = '300px';
      win.el.style.height = '200px';
      const result = app._test.getSnapEdges(win);
      const area = document.getElementById('window-area');
      return {
        x: result.x.sort((a, b) => a - b),
        y: result.y.sort((a, b) => a - b),
        areaW: area.clientWidth,
        areaH: area.clientHeight,
      };
    });
    // Should include 0 and areaW/areaH (area edges), but NOT the excluded window's edges
    expect(edges.x).toContain(0);
    expect(edges.x).toContain(edges.areaW);
    expect(edges.y).toContain(0);
    expect(edges.y).toContain(edges.areaH);
    // With only one window (excluded), should have exactly the area edges
    expect(edges.x).toEqual([0, edges.areaW]);
    expect(edges.y).toEqual([0, edges.areaH]);
  });

  test('getSnapEdges includes other window edges', async ({ page }) => {
    await uploadFile(page, '../test/sample2.psv');
    await waitForWindow(page, 'sample2');

    const edges = await page.evaluate(() => {
      const [w0, w1] = app._test.windows;
      w0.el.style.left = '100px';
      w0.el.style.top = '80px';
      w0.el.style.width = '300px';
      w0.el.style.height = '200px';
      w1.el.style.left = '500px';
      w1.el.style.top = '50px';
      w1.el.style.width = '400px';
      w1.el.style.height = '250px';
      // Get edges excluding window 0 — should include window 1's edges
      const result = app._test.getSnapEdges(w0);
      return {
        x: result.x.sort((a, b) => a - b),
        y: result.y.sort((a, b) => a - b),
      };
    });
    // Window 1 edges: left=500, right=900, top=50, bottom=300
    expect(edges.x).toContain(500);
    expect(edges.x).toContain(900);
    expect(edges.y).toContain(50);
    expect(edges.y).toContain(300);
  });

  test('getSnapEdges excludes minimized windows', async ({ page }) => {
    await uploadFile(page, '../test/sample2.psv');
    await waitForWindow(page, 'sample2');

    const edges = await page.evaluate(() => {
      const [w0, w1] = app._test.windows;
      w1.el.style.left = '500px';
      w1.el.style.top = '50px';
      w1.el.style.width = '400px';
      w1.el.style.height = '250px';
      w1.el.classList.add('minimized');
      const result = app._test.getSnapEdges(w0);
      const area = document.getElementById('window-area');
      return {
        x: result.x.sort((a, b) => a - b),
        areaW: area.clientWidth,
      };
    });
    // Minimized window's edges should NOT be included
    expect(edges.x).not.toContain(500);
    expect(edges.x).not.toContain(900);
    expect(edges.x).toEqual([0, edges.areaW]);
  });

  test('snapPosition snaps closer edge when both within threshold', async ({ page }) => {
    const result = await page.evaluate(() => {
      // Set up a single window so only area edges (0 and areaW) exist
      const win = app._test.windows[0];
      const area = document.getElementById('window-area');
      const areaW = area.clientWidth;
      const areaH = area.clientHeight;
      // Window at left=5, width=300 → left edge is 5 from 0, right edge is 305 from areaW
      // Left edge is closer, should snap left to 0
      const snap = app._test.snapPosition(win, 5, 5, 300, 200, areaW, areaH);
      return { left: snap.left, top: snap.top };
    });
    expect(result.left).toBe(0);
    expect(result.top).toBe(0);
  });
});

test.describe('Window Snap — Dialogs do not snap', () => {
  test('snapPosition does not snap a dialog window', async ({ page }) => {
    await openApp(page);
    // Ctrl+N opens the New Table prompt — a modal dialog (isDialog: true).
    await page.keyboard.press('Control+n');
    await expect(page.locator('.subwindow.dialog')).toHaveCount(1);

    const result = await page.evaluate(() => {
      const dlg = app._test.windows.find(w => w.isDialog);
      const area = document.getElementById('window-area');
      const areaW = area.clientWidth, areaH = area.clientHeight;
      // left/top=5 would normally snap to the area's 0 edges (within threshold).
      const snap = app._test.snapPosition(dlg, 5, 5, 300, 200, areaW, areaH);
      return { left: snap.left, top: snap.top, guidesX: snap.guidesX.length, guidesY: snap.guidesY.length };
    });
    // Position is left unsnapped (only clamped), and no guides are produced.
    expect(result.left).toBe(5);
    expect(result.top).toBe(5);
    expect(result.guidesX).toBe(0);
    expect(result.guidesY).toBe(0);
  });

  test('getSnapEdges excludes dialog windows (others do not snap to dialogs)', async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
    await page.keyboard.press('Control+n');
    await expect(page.locator('.subwindow.dialog')).toHaveCount(1);

    const edges = await page.evaluate(() => {
      const dataWin = app._test.windows.find(w => !w.isDialog);
      const dlg = app._test.windows.find(w => w.isDialog);
      dlg.el.style.left = '500px';
      dlg.el.style.top = '50px';
      dlg.el.style.width = '400px';
      dlg.el.style.height = '250px';
      const result = app._test.getSnapEdges(dataWin);
      const area = document.getElementById('window-area');
      return { x: result.x.sort((a, b) => a - b), areaW: area.clientWidth };
    });
    // Dialog edges (500, 900) must not appear — only the area edges remain.
    expect(edges.x).not.toContain(500);
    expect(edges.x).not.toContain(900);
    expect(edges.x).toEqual([0, edges.areaW]);
  });

  test('resizing a dialog does not snap to another window edge', async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');

    // Data window: right edge at 600 — a snap target for normal windows.
    await page.evaluate(() => {
      const dataWin = app._test.windows[0];
      dataWin.el.style.left = '0px';
      dataWin.el.style.top = '0px';
      dataWin.el.style.width = '600px';
      dataWin.el.style.height = '400px';
    });

    // Open a modal dialog and size it so its right edge sits just short of 600.
    await page.keyboard.press('Control+n');
    await expect(page.locator('.subwindow.dialog')).toHaveCount(1);
    await page.evaluate(() => {
      const dlg = app._test.windows.find(w => w.isDialog);
      dlg.el.style.left = '100px';
      dlg.el.style.top = '50px';
      dlg.el.style.width = '485px'; // right edge at 585
      dlg.el.style.height = '200px';
    });

    const handle = page.locator('.subwindow.dialog .resize-handle.rh-right');
    const box = await handle.boundingBox();
    const startX = box.x + box.width / 2;
    const startY = box.y + box.height / 2;
    await page.mouse.move(startX, startY);
    await page.mouse.down();
    // Drag right by 10px: right edge ~595, within SNAP_THRESHOLD (10) of 600.
    await page.mouse.move(startX + 10, startY, { steps: 10 });

    // No snap guides should appear while resizing a dialog.
    const guideCount = await page.evaluate(() => document.querySelectorAll('.snap-guide').length);
    expect(guideCount).toBe(0);

    await page.mouse.up();

    const rightEdge = await page.evaluate(() => {
      const dlg = app._test.windows.find(w => w.isDialog);
      return parseInt(dlg.el.style.left) + parseInt(dlg.el.style.width);
    });
    // A normal window would snap its right edge to 600. The dialog must not.
    expect(rightEdge).not.toBe(600);
    expect(rightEdge).toBeGreaterThanOrEqual(590);
    expect(rightEdge).toBeLessThan(600);
  });
});
