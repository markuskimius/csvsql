const path = require('path');

/**
 * Poll page.evaluate until predicate returns truthy, with timeout.
 */
async function pollUntil(page, fn, arg, { timeout = 10000, interval = 25 } = {}) {
  const deadline = Date.now() + timeout;
  while (Date.now() < deadline) {
    const result = arg !== undefined
      ? await page.evaluate(fn, arg)
      : await page.evaluate(fn);
    if (result) return result;
    await page.waitForTimeout(interval);
  }
  throw new Error(`pollUntil timed out after ${timeout}ms`);
}

/**
 * Navigate to the app with ?test=1 and wait for sql.js to initialize.
 */
async function openApp(page) {
  await page.goto('/?test=1');
  await pollUntil(page, () => window._appReady === true, undefined, { timeout: 15000 });
}

/**
 * Upload a file via the hidden #file-input element.
 */
async function uploadFile(page, filePath, { hasHeader = true } = {}) {
  const abs = path.resolve(__dirname, filePath);
  if (!hasHeader) {
    // Set shift-open flag so file loads without headers (columns become A, B, C, ...)
    await page.evaluate(() => { app._test._shiftOpen = true; });
  }
  const fileInput = page.locator('#file-input');
  await fileInput.setInputFiles(abs);
}

/**
 * Wait for a subwindow whose title contains the given substring.
 */
async function waitForWindow(page, titleSubstring) {
  await pollUntil(page, (sub) => {
    const titles = document.querySelectorAll('.subwindow .win-title');
    return [...titles].some(t => t.textContent.includes(sub));
  }, titleSubstring, { timeout: 10000 });
}

/**
 * Get table data from app internals.
 */
async function getTableData(page, tableName) {
  return page.evaluate((name) => {
    const t = app._test.tables[name];
    if (!t) return null;
    return {
      columns: t.columns,
      rows: t.rows.map(r => ({ ...r })),
      filename: t.filename,
      modified: t.modified,
    };
  }, tableName);
}

/**
 * Execute a SQL query via the console and wait for completion.
 * Waits until the status bar shows a final status (not "Running query...").
 */
async function executeSQL(page, sql) {
  await page.locator('#sql-input').fill(sql);
  await page.locator('#console-actions button:first-child').click();
  // Wait for the query to finish — status will change from "Running query..." to a result
  await pollUntil(page, () => {
    const el = document.getElementById('console-status');
    if (!el || !el.textContent) return null;
    // Still running
    if (el.textContent.startsWith('Running query')) return null;
    return el.textContent;
  }, undefined, { timeout: 15000 });
  return page.locator('#console-status').textContent();
}

/**
 * Count visible subwindows (not minimized).
 */
async function countVisibleWindows(page) {
  return page.evaluate(() => {
    return document.querySelectorAll('.subwindow:not(.minimized)').length;
  });
}

/**
 * Get the number of windows in the app's internal windows array.
 */
async function getWindowCount(page) {
  return page.evaluate(() => app._test.windows.length);
}

/**
 * Dispatch a touch event on an element at the given viewport coordinates.
 * Shared by the touch specs (touch-gestures, touch-drag-bridge, touch-long-press,
 * touch-dock, maximized-drag).
 */
async function dispatchTouch(page, elementHandle, type, clientX, clientY) {
  await elementHandle.evaluate((el, args) => {
    const touch = new Touch({
      identifier: 1,
      target: el,
      clientX: args.x,
      clientY: args.y,
      pageX: args.x,
      pageY: args.y,
      radiusX: 1,
      radiusY: 1,
      force: 1,
    });
    const list = args.type === 'touchend' || args.type === 'touchcancel' ? [] : [touch];
    const evt = new TouchEvent(args.type, {
      cancelable: true,
      bubbles: true,
      touches: list,
      targetTouches: list,
      changedTouches: [touch],
    });
    el.dispatchEvent(evt);
  }, { type, x: clientX, y: clientY });
}

/** Center viewport coordinates of a locator's bounding box. */
async function touchCenterOf(locator) {
  const box = await locator.boundingBox();
  return { x: box.x + box.width / 2, y: box.y + box.height / 2 };
}

module.exports = {
  openApp,
  uploadFile,
  waitForWindow,
  getTableData,
  executeSQL,
  countVisibleWindows,
  getWindowCount,
  dispatchTouch,
  touchCenterOf,
};
