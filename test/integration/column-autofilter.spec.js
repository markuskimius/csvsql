const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, getTableData } = require('../helpers');

test.describe('Column AutoFilter', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  test('filter icon is visible on column headers', async ({ page }) => {
    const filterBtns = page.locator('.data-table th:not(.row-num-header) .col-filter-btn');
    const count = await filterBtns.count();
    expect(count).toBe(3); // name, email, member_since
    // Icons should be visible (opacity > 0) without hover
    for (let i = 0; i < count; i++) {
      const opacity = await filterBtns.nth(i).evaluate(el => parseFloat(getComputedStyle(el).opacity));
      expect(opacity).toBeGreaterThan(0);
    }
  });

  test('clicking filter icon opens dropdown with unique values', async ({ page }) => {
    const th = page.locator('.data-table th:not(.row-num-header)').first();
    await th.locator('.col-filter-btn').click();
    await page.waitForTimeout(300);

    const dropdown = page.locator('.autofilter-dropdown');
    await expect(dropdown).toBeVisible();

    // Should have search, Select All, items, and buttons
    await expect(dropdown.locator('.autofilter-search')).toBeVisible();
    await expect(dropdown.locator('.autofilter-select-all')).toBeVisible();
    await expect(dropdown.locator('.autofilter-apply')).toBeVisible();
    await expect(dropdown.locator('.autofilter-clear')).toBeVisible();

    // Should have 10 unique name values
    const items = dropdown.locator('.autofilter-item');
    expect(await items.count()).toBe(10);
  });

  test('unchecking values and applying filters rows', async ({ page }) => {
    const th = page.locator('.data-table th:not(.row-num-header)').first();
    await th.locator('.col-filter-btn').click();
    await page.waitForTimeout(300);

    const dropdown = page.locator('.autofilter-dropdown');
    // Uncheck first two values
    const checkboxes = dropdown.locator('.autofilter-item input[type="checkbox"]');
    await checkboxes.nth(0).uncheck();
    await checkboxes.nth(1).uncheck();
    await dropdown.locator('.autofilter-apply').click();
    await page.waitForTimeout(300);

    // Should show 8 of 10 rows
    const rows = await page.locator('.data-table tbody tr:not(.virtual-pad)').count();
    expect(rows).toBe(8);
  });

  test('filtered column shows visual indicators', async ({ page }) => {
    const th = page.locator('.data-table th:not(.row-num-header)').first();
    await th.locator('.col-filter-btn').click();
    await page.waitForTimeout(300);

    // Uncheck one value and apply
    await page.locator('.autofilter-dropdown .autofilter-item input[type="checkbox"]').first().uncheck();
    await page.locator('.autofilter-dropdown .autofilter-apply').click();
    await page.waitForTimeout(300);

    // Header should have col-filtered class (green border)
    await expect(page.locator('.data-table th.col-filtered')).toHaveCount(1);
    // Filter button should have active class
    await expect(page.locator('.col-filter-btn.active')).toHaveCount(1);
  });

  test('Clear button removes column filter', async ({ page }) => {
    const th = page.locator('.data-table th:not(.row-num-header)').first();
    // Apply a filter first
    await th.locator('.col-filter-btn').click();
    await page.waitForTimeout(300);
    await page.locator('.autofilter-dropdown .autofilter-item input[type="checkbox"]').first().uncheck();
    await page.locator('.autofilter-dropdown .autofilter-apply').click();
    await page.waitForTimeout(300);
    expect(await page.locator('.data-table tbody tr:not(.virtual-pad)').count()).toBe(9);

    // Now clear
    await th.locator('.col-filter-btn').click();
    await page.waitForTimeout(300);
    await page.locator('.autofilter-dropdown .autofilter-clear').click();
    await page.waitForTimeout(300);

    // All 10 rows back
    expect(await page.locator('.data-table tbody tr:not(.virtual-pad)').count()).toBe(10);
    await expect(page.locator('.data-table th.col-filtered')).toHaveCount(0);
  });

  test('search box filters the value list', async ({ page }) => {
    const th = page.locator('.data-table th:not(.row-num-header)').first();
    await th.locator('.col-filter-btn').click();
    await page.waitForTimeout(300);

    await page.locator('.autofilter-search').fill('alice');
    await page.waitForTimeout(300);

    const items = page.locator('.autofilter-dropdown .autofilter-item');
    expect(await items.count()).toBe(1);
    const text = await items.first().textContent();
    expect(text).toContain('Alice');
  });

  test('Select All toggles all visible checkboxes', async ({ page }) => {
    const th = page.locator('.data-table th:not(.row-num-header)').first();
    await th.locator('.col-filter-btn').click();
    await page.waitForTimeout(300);

    const selectAll = page.locator('.autofilter-select-all input[type="checkbox"]');

    // Uncheck Select All
    await selectAll.uncheck();
    await page.waitForTimeout(100);

    // All items should be unchecked
    const checked = await page.locator('.autofilter-dropdown .autofilter-item input[type="checkbox"]:checked').count();
    expect(checked).toBe(0);

    // Re-check Select All
    await selectAll.check();
    await page.waitForTimeout(100);

    const checkedAfter = await page.locator('.autofilter-dropdown .autofilter-item input[type="checkbox"]:checked').count();
    expect(checkedAfter).toBe(10);
  });

  test('multiple column filters AND together', async ({ page }) => {
    // Filter name column to only Alice Johnson
    const nameTh = page.locator('.data-table th:not(.row-num-header)').nth(0);
    await nameTh.locator('.col-filter-btn').click();
    await page.waitForTimeout(300);

    // Uncheck all, then check only Alice Johnson
    await page.locator('.autofilter-select-all input[type="checkbox"]').uncheck();
    await page.waitForTimeout(100);
    await page.locator('.autofilter-dropdown .autofilter-item input[type="checkbox"]').first().check();
    await page.locator('.autofilter-dropdown .autofilter-apply').click();
    await page.waitForTimeout(300);
    expect(await page.locator('.data-table tbody tr:not(.virtual-pad)').count()).toBe(1);

    // Now also filter email column — the one visible email should be alice@example.com
    const emailTh = page.locator('.data-table th:not(.row-num-header)').nth(1);
    await emailTh.locator('.col-filter-btn').click();
    await page.waitForTimeout(300);

    // All 10 emails should be in the dropdown (unique values from all rows, not just filtered)
    const emailItems = page.locator('.autofilter-dropdown .autofilter-item');
    expect(await emailItems.count()).toBe(10);

    // Uncheck all, check only alice@example.com
    await page.locator('.autofilter-select-all input[type="checkbox"]').uncheck();
    await page.waitForTimeout(100);
    await page.locator('.autofilter-dropdown .autofilter-item input[type="checkbox"]').first().check();
    await page.locator('.autofilter-dropdown .autofilter-apply').click();
    await page.waitForTimeout(300);

    // Should still show 1 row (AND of both filters)
    expect(await page.locator('.data-table tbody tr:not(.virtual-pad)').count()).toBe(1);
    // Both columns should show col-filtered
    await expect(page.locator('.data-table th.col-filtered')).toHaveCount(2);
  });

  test('column filter combines with WHERE clause filter', async ({ page }) => {
    // Apply column filter: exclude first name
    const th = page.locator('.data-table th:not(.row-num-header)').first();
    await th.locator('.col-filter-btn').click();
    await page.waitForTimeout(300);
    await page.locator('.autofilter-dropdown .autofilter-item input[type="checkbox"]').first().uncheck();
    await page.locator('.autofilter-dropdown .autofilter-apply').click();
    await page.waitForTimeout(300);
    expect(await page.locator('.data-table tbody tr:not(.virtual-pad)').count()).toBe(9);

    // Now also apply a WHERE filter
    await page.locator('.subwindow .toolbar-mode .mode-filter').click();
    const filterInput = page.locator('.subwindow .filter-input');
    await filterInput.fill("[name] REGEXP '^[A-D]'");
    await page.waitForTimeout(500);

    // Should show intersection: names A-D minus the excluded name
    const rows = await page.locator('.data-table tbody tr:not(.virtual-pad)').count();
    expect(rows).toBeGreaterThan(0);
    expect(rows).toBeLessThan(9);
  });

  test('filter persists after sorting', async ({ page }) => {
    // Apply a column filter
    const th = page.locator('.data-table th:not(.row-num-header)').first();
    await th.locator('.col-filter-btn').click();
    await page.waitForTimeout(300);
    await page.locator('.autofilter-dropdown .autofilter-item input[type="checkbox"]').first().uncheck();
    await page.locator('.autofilter-dropdown .autofilter-apply').click();
    await page.waitForTimeout(300);
    expect(await page.locator('.data-table tbody tr:not(.virtual-pad)').count()).toBe(9);

    // Sort by clicking the sort badge
    await th.locator('.col-badge-sort').click();
    await page.waitForTimeout(300);

    // Filter should still be active
    expect(await page.locator('.data-table tbody tr:not(.virtual-pad)').count()).toBe(9);
    await expect(page.locator('.data-table th.col-filtered')).toHaveCount(1);
    await expect(page.locator('.data-table th.sorted')).toHaveCount(1);
  });

  test('dropdown closes on outside click', async ({ page }) => {
    const th = page.locator('.data-table th:not(.row-num-header)').first();
    await th.locator('.col-filter-btn').click();
    await page.waitForTimeout(300);
    await expect(page.locator('.autofilter-dropdown')).toBeVisible();

    // Click outside
    await page.mouse.click(100, 100);
    await page.waitForTimeout(200);
    await expect(page.locator('.autofilter-dropdown')).toHaveCount(0);
  });

  test('dropdown closes on Escape', async ({ page }) => {
    const th = page.locator('.data-table th:not(.row-num-header)').first();
    await th.locator('.col-filter-btn').click();
    await page.waitForTimeout(300);
    await expect(page.locator('.autofilter-dropdown')).toBeVisible();

    await page.keyboard.press('Escape');
    await page.waitForTimeout(200);
    await expect(page.locator('.autofilter-dropdown')).toHaveCount(0);
  });

  test('only one dropdown is open at a time', async ({ page }) => {
    const th1 = page.locator('.data-table th:not(.row-num-header)').nth(0);
    const th2 = page.locator('.data-table th:not(.row-num-header)').nth(1);

    await th1.locator('.col-filter-btn').click();
    await page.waitForTimeout(300);
    await expect(page.locator('.autofilter-dropdown')).toHaveCount(1);

    // Open a different column's filter
    await th2.locator('.col-filter-btn').click();
    await page.waitForTimeout(300);
    // Still only 1 dropdown
    await expect(page.locator('.autofilter-dropdown')).toHaveCount(1);
  });

  test('toggling same filter icon closes dropdown', async ({ page }) => {
    const th = page.locator('.data-table th:not(.row-num-header)').first();
    await th.locator('.col-filter-btn').click();
    await page.waitForTimeout(300);
    await expect(page.locator('.autofilter-dropdown')).toBeVisible();

    // Click same icon again
    await th.locator('.col-filter-btn').click();
    await page.waitForTimeout(200);
    await expect(page.locator('.autofilter-dropdown')).toHaveCount(0);
  });

  test('filter icon click does not trigger sort', async ({ page }) => {
    const th = page.locator('.data-table th:not(.row-num-header)').first();

    // Click the filter icon
    await th.locator('.col-filter-btn').click();
    await page.waitForTimeout(300);

    // Header should NOT be sorted
    await expect(th).not.toHaveClass(/sorted/);

    // Close dropdown
    await page.keyboard.press('Escape');
  });

  test('dropdown reopens with correct state after filter applied', async ({ page }) => {
    const th = page.locator('.data-table th:not(.row-num-header)').first();

    // Apply a filter (uncheck first two)
    await th.locator('.col-filter-btn').click();
    await page.waitForTimeout(300);
    const checkboxes = page.locator('.autofilter-dropdown .autofilter-item input[type="checkbox"]');
    await checkboxes.nth(0).uncheck();
    await checkboxes.nth(1).uncheck();
    await page.locator('.autofilter-dropdown .autofilter-apply').click();
    await page.waitForTimeout(300);

    // Reopen dropdown
    await th.locator('.col-filter-btn').click();
    await page.waitForTimeout(300);

    // First two should be unchecked, rest checked
    const cb = page.locator('.autofilter-dropdown .autofilter-item input[type="checkbox"]');
    expect(await cb.nth(0).isChecked()).toBe(false);
    expect(await cb.nth(1).isChecked()).toBe(false);
    expect(await cb.nth(2).isChecked()).toBe(true);
  });

  test('applying with all values checked removes the filter', async ({ page }) => {
    const th = page.locator('.data-table th:not(.row-num-header)').first();

    // Apply a filter first
    await th.locator('.col-filter-btn').click();
    await page.waitForTimeout(300);
    await page.locator('.autofilter-dropdown .autofilter-item input[type="checkbox"]').first().uncheck();
    await page.locator('.autofilter-dropdown .autofilter-apply').click();
    await page.waitForTimeout(300);
    await expect(page.locator('.data-table th.col-filtered')).toHaveCount(1);

    // Reopen and check all again
    await th.locator('.col-filter-btn').click();
    await page.waitForTimeout(300);
    await page.locator('.autofilter-select-all input[type="checkbox"]').check();
    await page.locator('.autofilter-dropdown .autofilter-apply').click();
    await page.waitForTimeout(300);

    // Filter should be removed (no col-filtered class)
    await expect(page.locator('.data-table th.col-filtered')).toHaveCount(0);
    expect(await page.locator('.data-table tbody tr:not(.virtual-pad)').count()).toBe(10);
  });

  test('filter icon uses SVG funnel and sort badge appears on sort', async ({ page }) => {
    const th = page.locator('.data-table th:not(.row-num-header)').first();

    // Filter icon should be an SVG funnel
    const hasSvg = await th.locator('.col-filter-btn svg').count();
    expect(hasSvg).toBe(1);

    // Sort the column
    await th.locator('.col-name').click();
    await page.waitForTimeout(300);

    // Sort badge should appear
    await expect(th.locator('.col-badge-sort')).toHaveCount(1);
  });

  test('sort badge is inline in header cell next to filter button', async ({ page }) => {
    const th = page.locator('.data-table th:not(.row-num-header)').first();

    // Sort the column so badge is filled
    await th.locator('.col-name').click();
    await page.waitForTimeout(300);

    const sortBadge = th.locator('.th-inner .col-badge-sort');
    const filterBtn = th.locator('.th-inner .col-filter-btn');
    expect(await sortBadge.count()).toBe(1);
    expect(await filterBtn.count()).toBe(1);

    // Sort badge should be to the left of filter button
    const sortBox = await sortBadge.boundingBox();
    const filterBox = await filterBtn.boundingBox();
    expect(sortBox.x + sortBox.width).toBeLessThanOrEqual(filterBox.x + 2);
  });

  test('Filtered chip not shown when no filters active', async ({ page }) => {
    await expect(page.locator('.status-chip-filter.chip-inactive')).toHaveCount(1);
    const statusText = await page.locator('.status-left').textContent();
    expect(statusText).toBe('10 of 10 rows');
  });

  test('Filtered chip appears with column autofilter', async ({ page }) => {
    const th = page.locator('.data-table th:not(.row-num-header)').first();
    await th.locator('.col-filter-btn').click();
    await page.waitForTimeout(300);
    await page.locator('.autofilter-dropdown .autofilter-item input[type="checkbox"]').first().uncheck();
    await page.locator('.autofilter-dropdown .autofilter-apply').click();
    await page.waitForTimeout(300);

    await expect(page.locator('.status-chip-filter:not(.chip-inactive)')).toHaveCount(1);
    expect(await page.locator('.status-chip-filter').textContent()).toBe('Filtered');
  });

  test('Filtered chip appears with WHERE filter', async ({ page }) => {
    await page.locator('.subwindow .toolbar-mode .mode-filter').click();
    const filterInput = page.locator('.subwindow .filter-input');
    await filterInput.fill("name LIKE '%Alice%'");
    await page.waitForTimeout(500);

    await expect(page.locator('.status-chip-filter:not(.chip-inactive)')).toHaveCount(1);
  });

  test('Filtered chip clears column autofilters when clicked', async ({ page }) => {
    // Apply column filter on two columns
    const nameTh = page.locator('.data-table th:not(.row-num-header)').nth(0);
    await nameTh.locator('.col-filter-btn').click();
    await page.waitForTimeout(300);
    await page.locator('.autofilter-dropdown .autofilter-item input[type="checkbox"]').first().uncheck();
    await page.locator('.autofilter-dropdown .autofilter-apply').click();
    await page.waitForTimeout(300);

    const emailTh = page.locator('.data-table th:not(.row-num-header)').nth(1);
    await emailTh.locator('.col-filter-btn').click();
    await page.waitForTimeout(300);
    await page.locator('.autofilter-dropdown .autofilter-item input[type="checkbox"]').first().uncheck();
    await page.locator('.autofilter-dropdown .autofilter-apply').click();
    await page.waitForTimeout(300);

    await expect(page.locator('.data-table th.col-filtered')).toHaveCount(2);

    // Click Filtered chip
    await page.locator('.status-chip-filter').click();
    await page.waitForTimeout(300);

    // All rows back, no filtered indicators, chip inactive
    expect(await page.locator('.data-table tbody tr:not(.virtual-pad)').count()).toBe(10);
    await expect(page.locator('.data-table th.col-filtered')).toHaveCount(0);
    await expect(page.locator('.col-filter-btn.active')).toHaveCount(0);
    await expect(page.locator('.status-chip-filter.chip-inactive')).toHaveCount(1);
  });

  test('Filtered chip clears WHERE filter text when clicked', async ({ page }) => {
    await page.locator('.subwindow .toolbar-mode .mode-filter').click();
    const filterInput = page.locator('.subwindow .filter-input');
    await filterInput.fill("name LIKE '%Alice%'");
    await page.waitForTimeout(500);
    expect(await page.locator('.data-table tbody tr:not(.virtual-pad)').count()).toBe(1);

    await page.locator('.status-chip-filter').click();
    await page.waitForTimeout(300);

    expect(await filterInput.inputValue()).toBe('');
    expect(await page.locator('.data-table tbody tr:not(.virtual-pad)').count()).toBe(10);
    await expect(page.locator('.status-chip-filter.chip-inactive')).toHaveCount(1);
  });

  test('Filtered chip clears both column autofilters and WHERE at once', async ({ page }) => {
    // Apply column filter
    const th = page.locator('.data-table th:not(.row-num-header)').first();
    await th.locator('.col-filter-btn').click();
    await page.waitForTimeout(300);
    await page.locator('.autofilter-select-all input[type="checkbox"]').uncheck();
    await page.waitForTimeout(100);
    await page.locator('.autofilter-dropdown .autofilter-item input[type="checkbox"]').first().check();
    await page.locator('.autofilter-dropdown .autofilter-apply').click();
    await page.waitForTimeout(300);

    // Also apply WHERE filter
    await page.locator('.subwindow .toolbar-mode .mode-filter').click();
    const filterInput = page.locator('.subwindow .filter-input');
    await filterInput.fill("[email] LIKE '%alice%'");
    await page.waitForTimeout(500);

    expect(await page.locator('.data-table tbody tr:not(.virtual-pad)').count()).toBe(1);
    await expect(page.locator('.data-table th.col-filtered')).toHaveCount(1);

    // Click Filtered chip — both should be gone
    await page.locator('.status-chip-filter').click();
    await page.waitForTimeout(300);

    expect(await page.locator('.data-table tbody tr:not(.virtual-pad)').count()).toBe(10);
    await expect(page.locator('.data-table th.col-filtered')).toHaveCount(0);
    expect(await filterInput.inputValue()).toBe('');
    await expect(page.locator('.status-chip-filter.chip-inactive')).toHaveCount(1);
  });

  test('no Clear Filters link in status bar', async ({ page }) => {
    // Apply a filter and verify the old Clear Filters link is gone
    const th = page.locator('.data-table th:not(.row-num-header)').first();
    await th.locator('.col-filter-btn').click();
    await page.waitForTimeout(300);
    await page.locator('.autofilter-dropdown .autofilter-item input[type="checkbox"]').first().uncheck();
    await page.locator('.autofilter-dropdown .autofilter-apply').click();
    await page.waitForTimeout(300);

    await expect(page.locator('.status-clear-filters')).toHaveCount(0);
    const statusText = await page.locator('.status-left').textContent();
    expect(statusText).not.toContain('Clear Filters');
  });

  test('status bar shows filtered row count without trailing dash', async ({ page }) => {
    await page.locator('.subwindow .toolbar-mode .mode-filter').click();
    const filterInput = page.locator('.subwindow .filter-input');
    await filterInput.fill("name LIKE '%Alice%'");
    await page.waitForTimeout(500);

    const statusText = await page.locator('.status-left').textContent();
    expect(statusText).toBe('1 of 10 rows');
    expect(statusText).not.toContain('—');
  });

  test('dropdown is right-aligned to column header', async ({ page }) => {
    const th = page.locator('.data-table th:not(.row-num-header)').first();
    await th.locator('.col-filter-btn').click();
    await page.waitForTimeout(300);

    const thRect = await th.boundingBox();
    const ddRect = await page.locator('.autofilter-dropdown').boundingBox();

    // Dropdown right edge should be close to header right edge (within a few px)
    expect(Math.abs((ddRect.x + ddRect.width) - (thRect.x + thRect.width))).toBeLessThan(5);
  });
});
