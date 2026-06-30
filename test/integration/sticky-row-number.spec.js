const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow } = require('../helpers');

test.describe('Sticky Row Number Column', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
  });

  test('row-num header has position sticky and left 0', async ({ page }) => {
    await uploadFile(page, 'sample1.csv');
    await waitForWindow(page, 'sample1');

    const styles = await page.$eval('th.row-num-header', el => {
      const cs = getComputedStyle(el);
      return { position: cs.position, left: cs.left };
    });
    expect(styles.position).toBe('sticky');
    expect(styles.left).toBe('0px');
  });

  test('row-num data cells have position sticky and left 0', async ({ page }) => {
    await uploadFile(page, 'sample1.csv');
    await waitForWindow(page, 'sample1');

    const styles = await page.$eval('td.row-num', el => {
      const cs = getComputedStyle(el);
      return { position: cs.position, left: cs.left };
    });
    expect(styles.position).toBe('sticky');
    expect(styles.left).toBe('0px');
  });

  test('row-num header z-index is higher than regular headers', async ({ page }) => {
    await uploadFile(page, 'sample1.csv');
    await waitForWindow(page, 'sample1');

    const zIndexes = await page.evaluate(() => {
      const rowNumHeader = document.querySelector('th.row-num-header');
      const dataHeader = document.querySelector('th:not(.row-num-header)');
      return {
        rowNumHeader: parseInt(getComputedStyle(rowNumHeader).zIndex, 10),
        dataHeader: parseInt(getComputedStyle(dataHeader).zIndex, 10),
      };
    });
    expect(zIndexes.rowNumHeader).toBeGreaterThan(zIndexes.dataHeader);
  });

  test('row-num data cell z-index is >= regular header z-index', async ({ page }) => {
    await uploadFile(page, 'sample1.csv');
    await waitForWindow(page, 'sample1');

    const zIndexes = await page.evaluate(() => {
      const rowNumCell = document.querySelector('td.row-num');
      const dataHeader = document.querySelector('th:not(.row-num-header)');
      return {
        rowNumCell: parseInt(getComputedStyle(rowNumCell).zIndex, 10),
        dataHeader: parseInt(getComputedStyle(dataHeader).zIndex, 10),
      };
    });
    expect(zIndexes.rowNumCell).toBeGreaterThanOrEqual(zIndexes.dataHeader);
  });

  test('row-num header has opaque background', async ({ page }) => {
    await uploadFile(page, 'sample1.csv');
    await waitForWindow(page, 'sample1');

    const bg = await page.$eval('th.row-num-header', el =>
      getComputedStyle(el).backgroundColor);
    expect(bg).not.toContain('rgba');
  });

  test('row-num data cells have opaque background', async ({ page }) => {
    await uploadFile(page, 'sample1.csv');
    await waitForWindow(page, 'sample1');

    const bg = await page.$eval('td.row-num', el =>
      getComputedStyle(el).backgroundColor);
    expect(bg).not.toContain('rgba');
  });

  test('table uses border-collapse separate for sticky compatibility', async ({ page }) => {
    await uploadFile(page, 'sample1.csv');
    await waitForWindow(page, 'sample1');

    const borderCollapse = await page.$eval('.data-table', el =>
      getComputedStyle(el).borderCollapse);
    expect(borderCollapse).toBe('separate');
  });

  test('table has zero border-spacing', async ({ page }) => {
    await uploadFile(page, 'sample1.csv');
    await waitForWindow(page, 'sample1');

    const borderSpacing = await page.$eval('.data-table', el =>
      getComputedStyle(el).borderSpacing);
    expect(['0px', '0px 0px']).toContain(borderSpacing);
  });

  test('row-num column stays fixed when scrolling right', async ({ page }) => {
    // Create a wide table that requires horizontal scrolling
    const cols = Array.from({ length: 20 }, (_, i) => `Col${i + 1}`);
    const rows = Array.from({ length: 5 }, (_, r) =>
      cols.map((_, c) => `R${r + 1}C${c + 1}`).join(',')
    );
    const csv = cols.join(',') + '\n' + rows.join('\n');
    const fileInput = page.locator('#file-input');
    await fileInput.setInputFiles({
      name: 'wide.csv',
      mimeType: 'text/csv',
      buffer: Buffer.from(csv),
    });
    await waitForWindow(page, 'wide');
    await page.waitForTimeout(200);

    // Record row-num header position before scrolling
    const container = page.locator('.table-container');
    const beforeLeft = await page.$eval('th.row-num-header', el => {
      return el.getBoundingClientRect().left;
    });

    // Scroll the table container to the right
    await container.evaluate(el => { el.scrollLeft = 300; });
    await page.waitForTimeout(100);

    // Row-num header should remain at the same screen position
    const afterLeft = await page.$eval('th.row-num-header', el => {
      return el.getBoundingClientRect().left;
    });
    expect(afterLeft).toBe(beforeLeft);
  });

  test('row-num data cells stay fixed when scrolling right', async ({ page }) => {
    const cols = Array.from({ length: 20 }, (_, i) => `Col${i + 1}`);
    const rows = Array.from({ length: 5 }, (_, r) =>
      cols.map((_, c) => `R${r + 1}C${c + 1}`).join(',')
    );
    const csv = cols.join(',') + '\n' + rows.join('\n');
    const fileInput = page.locator('#file-input');
    await fileInput.setInputFiles({
      name: 'wide.csv',
      mimeType: 'text/csv',
      buffer: Buffer.from(csv),
    });
    await waitForWindow(page, 'wide');
    await page.waitForTimeout(200);

    const container = page.locator('.table-container');
    const beforeLeft = await page.$eval('td.row-num', el => {
      return el.getBoundingClientRect().left;
    });

    await container.evaluate(el => { el.scrollLeft = 300; });
    await page.waitForTimeout(100);

    const afterLeft = await page.$eval('td.row-num', el => {
      return el.getBoundingClientRect().left;
    });
    expect(afterLeft).toBe(beforeLeft);
  });

  test('data columns scroll while row-num stays pinned', async ({ page }) => {
    const cols = Array.from({ length: 20 }, (_, i) => `Col${i + 1}`);
    const rows = Array.from({ length: 5 }, (_, r) =>
      cols.map((_, c) => `R${r + 1}C${c + 1}`).join(',')
    );
    const csv = cols.join(',') + '\n' + rows.join('\n');
    const fileInput = page.locator('#file-input');
    await fileInput.setInputFiles({
      name: 'wide.csv',
      mimeType: 'text/csv',
      buffer: Buffer.from(csv),
    });
    await waitForWindow(page, 'wide');
    await page.waitForTimeout(200);

    const container = page.locator('.table-container');

    // Get first data header position before scroll
    const beforeDataLeft = await page.$eval('th:not(.row-num-header)', el => {
      return el.getBoundingClientRect().left;
    });

    await container.evaluate(el => { el.scrollLeft = 300; });
    await page.waitForTimeout(100);

    // First data header should have moved left
    const afterDataLeft = await page.$eval('th:not(.row-num-header)', el => {
      return el.getBoundingClientRect().left;
    });
    expect(afterDataLeft).toBeLessThan(beforeDataLeft);
  });

  test('first column has left border', async ({ page }) => {
    await uploadFile(page, 'sample1.csv');
    await waitForWindow(page, 'sample1');

    const borderLeft = await page.$eval('th.row-num-header', el =>
      getComputedStyle(el).borderLeftWidth);
    expect(borderLeft).toBe('1px');
  });

  test('non-first-column cells have right and bottom borders', async ({ page }) => {
    await uploadFile(page, 'sample1.csv');
    await waitForWindow(page, 'sample1');

    const borders = await page.$eval('th:not(.row-num-header)', el => {
      const cs = getComputedStyle(el);
      return {
        right: cs.borderRightWidth,
        bottom: cs.borderBottomWidth,
      };
    });
    expect(borders.right).toBe('1px');
    expect(borders.bottom).toBe('1px');
  });

  test('header row has top border', async ({ page }) => {
    await uploadFile(page, 'sample1.csv');
    await waitForWindow(page, 'sample1');

    const borderTop = await page.$eval('thead th', el =>
      getComputedStyle(el).borderTopWidth);
    expect(borderTop).toBe('1px');
  });

  test('row-highlight on row-num cells uses opaque background', async ({ page }) => {
    await uploadFile(page, 'sample1.csv');
    await waitForWindow(page, 'sample1');

    // Click a data cell to trigger row highlight
    await page.click('td.data-cell');
    await page.waitForTimeout(100);

    // The row with the highlight should have the row-num cell with opaque bg
    const bg = await page.$eval('tr.row-highlight td.row-num', el =>
      getComputedStyle(el).backgroundColor);
    expect(bg).not.toContain('rgba');
  });
});
