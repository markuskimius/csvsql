const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, executeSQL } = require('../helpers');

// Quick-search debounce is 150ms, filter debounce 200ms — wait a bit longer.
const DEBOUNCE = 350;

// sample1.csv: name, email, member_since — 10 rows.
// "davi" matches David Brown, david@example.com, and Eve Davis (3 cells).
test.describe('Quick search (toolbar Search mode)', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  const input = (page) => page.locator('.subwindow .filter-input');
  const countEl = (page) => page.locator('.subwindow .search-count');

  async function enterFilterMode(page) {
    await page.locator('.subwindow .toolbar-mode .mode-filter').click();
  }

  test('toolbar defaults to Search mode, listed before Filter', async ({ page }) => {
    await expect(page.locator('.subwindow .toolbar-mode .mode-search')).toHaveClass(/active/);
    await expect(input(page)).toHaveAttribute('placeholder', /Quick search/);
    // Search is the first segment in the toggle
    await expect(page.locator('.subwindow .toolbar-mode button').first()).toHaveText('Search');
    // Filter mode still reachable with the WHERE placeholder
    await enterFilterMode(page);
    await expect(page.locator('.subwindow .toolbar-mode .mode-filter')).toHaveClass(/active/);
    await expect(input(page)).toHaveAttribute('placeholder', /WHERE clause/);
    await expect(countEl(page)).toBeHidden();
  });

  test('search mode highlights matches without filtering rows', async ({ page }) => {
    await input(page).fill('alice');
    await page.waitForTimeout(DEBOUNCE);
    await expect(countEl(page)).toHaveText('1 of 2');
    await expect(page.locator('.subwindow table tbody td.cell-find-match')).toHaveCount(2);
    await expect(page.locator('.subwindow table tbody td.cell-find-current')).toHaveCount(1);
    // Rows are not filtered
    await expect(page.locator('.subwindow .win-statusbar .status-left')).toContainText('10 of 10 rows');
  });

  test('Enter enters the grid at the first match; Tab/Shift+Tab step; Esc returns to the box', async ({ page }) => {
    const current = page.locator('.subwindow table tbody td.cell-find-current');
    await input(page).fill('davi');
    await page.waitForTimeout(DEBOUNCE);
    await expect(countEl(page)).toHaveText('1 of 3');
    // First Enter moves the cursor into the first match
    await input(page).press('Enter');
    await expect(current).toBeFocused();
    await expect(current).toHaveText('David Brown');
    // Tab / Shift+Tab step through matches
    await page.keyboard.press('Tab');
    await expect(current).toHaveText('david@example.com');
    await expect(countEl(page)).toHaveText('2 of 3');
    await page.keyboard.press('Tab');
    await expect(current).toHaveText('Eve Davis');
    // Stepping past the last match wraps to the first, with counter + toast feedback
    await page.keyboard.press('Tab');
    await expect(countEl(page)).toHaveText('1 of 3 — wrapped');
    await expect(countEl(page)).toHaveClass(/wrapped/);
    await expect(page.locator('.toast', { hasText: 'wrapped to the first match' })).toBeVisible();
    // Shift+Tab from the first match wraps back to the last
    await page.keyboard.press('Shift+Tab');
    await expect(countEl(page)).toHaveText('3 of 3 — wrapped');
    // A normal step clears the wrap flag
    await page.keyboard.press('Shift+Tab');
    await expect(countEl(page)).toHaveText('2 of 3');
    await expect(countEl(page)).not.toHaveClass(/wrapped/);
    // Escape stops navigating and returns the cursor to the search box
    await page.keyboard.press('Escape');
    await expect(input(page)).toBeFocused();
    await expect(page.locator('.subwindow table tbody td.cell-find-match')).toHaveCount(3);
    // Enter again resumes at the current match
    await input(page).press('Enter');
    await expect(current).toBeFocused();
    await expect(current).toHaveText('david@example.com');
  });

  test('Tab in the search input also enters the grid at the first match', async ({ page }) => {
    await input(page).fill('alice');
    await page.waitForTimeout(DEBOUNCE);
    await input(page).press('Tab');
    await expect(page.locator('.subwindow table tbody td.cell-find-current')).toBeFocused();
    await expect(page.locator('.subwindow table tbody td.cell-find-current')).toHaveText('Alice Johnson');
  });

  test('edit and selection shortcuts work during search navigation and end it', async ({ page }) => {
    const current = page.locator('.subwindow table tbody td.cell-find-current');
    await input(page).fill('davi');
    await page.waitForTimeout(DEBOUNCE);
    await input(page).press('Enter');
    await page.keyboard.press('Tab'); // 2nd match: david@example.com
    await expect(current).toHaveText('david@example.com');
    // "i" enters edit mode on the matched cell (and ends navigation)
    await page.keyboard.press('i');
    await expect(current).toHaveAttribute('contenteditable', 'true');
    await page.keyboard.press('Escape'); // revert edit, stay on the cell
    await expect(current).not.toHaveAttribute('contenteditable', 'true');
    // Navigation is over: Tab moves to the adjacent cell, not the next match
    await page.keyboard.press('Tab');
    const focused = page.locator('.subwindow table tbody td.data-cell:focus');
    await expect(focused).toHaveText('1780862826.123456730');
  });

  test('arrow keys move the selection normally and end search navigation', async ({ page }) => {
    await input(page).fill('davi');
    await page.waitForTimeout(DEBOUNCE);
    await input(page).press('Enter'); // on David Brown
    // Shift+arrow extends the selection (nav ends)
    await page.keyboard.press('Shift+ArrowDown');
    await expect(page.locator('.subwindow table tbody td.cell-selected')).toHaveCount(2);
    // Tab now moves to the adjacent cell instead of the next match
    await page.keyboard.press('Tab');
    const focused = page.locator('.subwindow table tbody td.data-cell:focus');
    await expect(focused).toHaveText('eve@example.com');
  });

  test('Escape returns focus to the current match and n/p navigate from the grid', async ({ page }) => {
    await input(page).fill('davi');
    await page.waitForTimeout(DEBOUNCE);
    await input(page).press('Escape');
    // Focus lands on the current (first) match: David Brown
    await expect(page.locator('.subwindow table tbody td.cell-find-current')).toBeFocused();
    await expect(page.locator('.subwindow table tbody td.cell-find-current')).toHaveText('David Brown');
    await page.keyboard.press('n');
    await expect(page.locator('.subwindow table tbody td.cell-find-current')).toHaveText('david@example.com');
    await page.keyboard.press('n');
    await expect(page.locator('.subwindow table tbody td.cell-find-current')).toHaveText('Eve Davis');
    await page.keyboard.press('p');
    await expect(page.locator('.subwindow table tbody td.cell-find-current')).toHaveText('david@example.com');
    // Wrap backward from the first match shows a toast
    await page.keyboard.press('p');
    await page.keyboard.press('p');
    await expect(page.locator('.subwindow table tbody td.cell-find-current')).toHaveText('Eve Davis');
    await expect(page.locator('.toast', { hasText: 'wrapped to the last match' })).toBeVisible();
  });

  test('/ from Filter mode opens search temporarily; Escape reverts keeping highlights', async ({ page }) => {
    await enterFilterMode(page);
    await page.locator('.subwindow table tbody td.data-cell').first().click();
    await page.keyboard.press('/');
    await expect(page.locator('.subwindow .toolbar-mode .mode-search')).toHaveClass(/active/);
    await expect(input(page)).toBeFocused();
    await input(page).fill('davi');
    await page.waitForTimeout(DEBOUNCE);
    await input(page).press('Escape');
    // Toggle flips back to Filter, but highlights and n/p stay active
    await expect(page.locator('.subwindow .toolbar-mode .mode-filter')).toHaveClass(/active/);
    await expect(input(page)).toHaveAttribute('placeholder', /WHERE clause/);
    await expect(page.locator('.subwindow table tbody td.cell-find-match')).toHaveCount(3);
    await page.keyboard.press('n');
    await expect(page.locator('.subwindow table tbody td.cell-find-current')).toHaveText('david@example.com');
  });

  test('/ in default Search mode just focuses the input, no revert on Escape', async ({ page }) => {
    await page.locator('.subwindow table tbody td.data-cell').first().click();
    await page.keyboard.press('/');
    await expect(input(page)).toBeFocused();
    await input(page).fill('alice');
    await page.waitForTimeout(DEBOUNCE);
    await input(page).press('Escape');
    // Still in Search mode — it was not a temporary switch
    await expect(page.locator('.subwindow .toolbar-mode .mode-search')).toHaveClass(/active/);
    await expect(page.locator('.subwindow table tbody td.cell-find-match')).toHaveCount(2);
  });

  test('Escape on a cell clears the search highlights before the selection', async ({ page }) => {
    await input(page).fill('alice');
    await page.waitForTimeout(DEBOUNCE);
    await input(page).press('Escape'); // back to the grid, highlights active
    await expect(page.locator('.subwindow table tbody td.cell-find-match')).toHaveCount(2);
    await page.keyboard.press('Escape'); // 1st Escape: clear search
    await expect(page.locator('.subwindow table tbody td.cell-find-match')).toHaveCount(0);
    await expect(page.locator('.subwindow table tbody td.cell-selected')).toHaveCount(1);
    await page.keyboard.press('Escape'); // 2nd Escape: clear selection
    await expect(page.locator('.subwindow table tbody td.cell-selected')).toHaveCount(0);
  });

  test('search respects the WHERE filter and both inputs keep their text across toggles', async ({ page }) => {
    await enterFilterMode(page);
    await input(page).fill("name LIKE 'A%'");
    await page.waitForTimeout(DEBOUNCE);
    await expect(page.locator('.subwindow .win-statusbar .status-left')).toContainText('1 of 10 rows');
    await page.locator('.subwindow .toolbar-mode .mode-search').click();
    // Rows stay filtered; the input swaps to the (empty) search text
    await expect(page.locator('.subwindow .win-statusbar .status-left')).toContainText('1 of 10 rows');
    await expect(input(page)).toHaveValue('');
    await input(page).fill('example.com');
    await page.waitForTimeout(DEBOUNCE);
    // Only the displayed (filtered) rows are searched
    await expect(countEl(page)).toHaveText('1 of 1');
    // Toggling back restores the WHERE text
    await enterFilterMode(page);
    await expect(input(page)).toHaveValue("name LIKE 'A%'");
  });

  test('Clear button appears with text and clears search or filter per mode', async ({ page }) => {
    const clearBtn = page.locator('.subwindow .filter-clear');
    await expect(clearBtn).toBeHidden();
    // Search mode: clears the text and highlights
    await input(page).fill('davi');
    await page.waitForTimeout(DEBOUNCE);
    await expect(clearBtn).toBeVisible();
    await expect(page.locator('.subwindow table tbody td.cell-find-match')).toHaveCount(3);
    await clearBtn.click();
    await expect(input(page)).toHaveValue('');
    await expect(clearBtn).toBeHidden();
    await expect(page.locator('.subwindow table tbody td.cell-find-match')).toHaveCount(0);
    await expect(input(page)).toBeFocused();
    // Filter mode: clears the WHERE clause and restores all rows
    await enterFilterMode(page);
    await expect(clearBtn).toBeHidden();
    await input(page).fill("name LIKE 'A%'");
    await page.waitForTimeout(DEBOUNCE);
    await expect(page.locator('.subwindow .win-statusbar .status-left')).toContainText('1 of 10 rows');
    await expect(clearBtn).toBeVisible();
    await clearBtn.click();
    await expect(input(page)).toHaveValue('');
    await expect(clearBtn).toBeHidden();
    await expect(page.locator('.subwindow .win-statusbar .status-left')).toContainText('10 of 10 rows');
  });

  test('Clear works in a /-opened temporary search; Escape still reverts to Filter', async ({ page }) => {
    const clearBtn = page.locator('.subwindow .filter-clear');
    await enterFilterMode(page);
    await input(page).fill("name LIKE 'A%'");
    await page.waitForTimeout(DEBOUNCE);
    // "/" from the grid opens a temporary search
    await page.locator('.subwindow table tbody td.data-cell').first().click();
    await page.keyboard.press('/');
    await expect(page.locator('.subwindow .toolbar-mode .mode-search')).toHaveClass(/active/);
    await input(page).fill('example.com');  // matches within the one filtered row
    await page.waitForTimeout(DEBOUNCE);
    await expect(page.locator('.subwindow table tbody td.cell-find-match')).toHaveCount(1);
    // Clear empties the search but keeps the temporary Search mode
    await clearBtn.click();
    await expect(input(page)).toHaveValue('');
    await expect(page.locator('.subwindow table tbody td.cell-find-match')).toHaveCount(0);
    await expect(page.locator('.subwindow .toolbar-mode .mode-search')).toHaveClass(/active/);
    // Escape still flips back to Filter with the WHERE text intact
    await input(page).press('Escape');
    await expect(page.locator('.subwindow .toolbar-mode .mode-filter')).toHaveClass(/active/);
    await expect(input(page)).toHaveValue("name LIKE 'A%'");
    await expect(clearBtn).toBeVisible();
  });

  test('Clear button visibility follows each mode\'s own text across toggles', async ({ page }) => {
    const clearBtn = page.locator('.subwindow .filter-clear');
    await input(page).fill('davi');
    await page.waitForTimeout(DEBOUNCE);
    await expect(clearBtn).toBeVisible();
    // Filter mode has no text — button hides
    await enterFilterMode(page);
    await expect(clearBtn).toBeHidden();
    // Back to Search — the retained search text brings it back
    await page.locator('.subwindow .toolbar-mode .mode-search').click();
    await expect(input(page)).toHaveValue('davi');
    await expect(clearBtn).toBeVisible();
  });

  test('Clear removes a filter SQL error state', async ({ page }) => {
    const clearBtn = page.locator('.subwindow .filter-clear');
    await enterFilterMode(page);
    await input(page).fill('name LIKE');  // incomplete WHERE clause
    await page.waitForTimeout(DEBOUNCE);
    await expect(input(page)).toHaveClass(/filter-error/);
    await clearBtn.click();
    await expect(input(page)).not.toHaveClass(/filter-error/);
    await expect(input(page)).toHaveValue('');
    await expect(page.locator('.subwindow .win-statusbar .status-left')).toContainText('10 of 10 rows');
  });

  test('quick-search matches recompute when the filter changes', async ({ page }) => {
    await input(page).fill('example.com');
    await page.waitForTimeout(DEBOUNCE);
    await expect(countEl(page)).toHaveText('1 of 10');
    // Apply a WHERE filter from the other mode; matches shrink to displayed rows
    await enterFilterMode(page);
    await input(page).fill("name LIKE 'A%'");
    await page.waitForTimeout(DEBOUNCE);
    await expect(page.locator('.subwindow table tbody td.cell-find-match')).toHaveCount(1);
  });

  test('/ works with no cell focused while the table window is active', async ({ page }) => {
    // Focus is on the titlebar area, not a data cell
    await page.locator('.subwindow .win-titlebar').click();
    await page.keyboard.press('/');
    await expect(input(page)).toBeFocused();
    await input(page).fill('alice');
    await page.waitForTimeout(DEBOUNCE);
    await expect(countEl(page)).toHaveText('1 of 2');
  });

  test('/ does not steal the keystroke from other inputs', async ({ page }) => {
    const sqlInput = page.locator('#sql-input');
    await sqlInput.click();
    await sqlInput.pressSequentially('a/b');
    await expect(sqlInput).toHaveValue('a/b');
    await expect(input(page)).not.toBeFocused();
  });

  test('search matches formatted values when formatting is enabled, raw when suspended', async ({ page }) => {
    await page.evaluate(() => {
      const cfg = {
        name: 'Redact',
        table: '.*',
        columns: [{ match: '^name$', display: "value == 'Alice Johnson' ? 'REDACTED' : value" }],
      };
      const errors = app._test.validatePlugin(cfg);
      if (errors.length) throw new Error(errors.join(', '));
      cfg._compiled = app._test.compilePlugin(cfg);
      app._test.plugins.push(cfg);
      app._test.rebuildTransformCache();
      app._test.rebuildTable(app._test.windows[0]);
    });
    // The formatted display value matches
    await input(page).fill('redacted');
    await page.waitForTimeout(DEBOUNCE);
    await expect(countEl(page)).toHaveText('1 of 1');
    // The raw value behind a transformed cell does not match — only the
    // untransformed email cell still contains "alice"
    await input(page).fill('alice');
    await page.waitForTimeout(DEBOUNCE);
    await expect(countEl(page)).toHaveText('1 of 1');
    // Suspending formatting via the Formatted chip switches back to raw
    // matching — "Alice Johnson" matches again (current match stays on email)
    await page.locator('.status-chip-format').click();
    await expect(countEl(page)).toHaveText('2 of 2');
    await input(page).fill('redacted');
    await page.waitForTimeout(DEBOUNCE);
    await expect(countEl(page)).toHaveText('No matches');
  });

  test('header cells are searched and navigable', async ({ page }) => {
    // 'member' only matches the member_since column header
    await input(page).fill('member');
    await page.waitForTimeout(DEBOUNCE);
    await expect(countEl(page)).toHaveText('1 of 1');
    await expect(page.locator('.subwindow table thead th.cell-find-match')).toHaveCount(1);
    await input(page).press('Enter');
    const th = page.locator('.subwindow table thead th.cell-find-current');
    await expect(th).toBeFocused();
    await expect(th).toContainText('member_since');
    // Escape returns to the search box
    await page.keyboard.press('Escape');
    await expect(input(page)).toBeFocused();
  });

  test('navigation steps from header matches into data matches and back', async ({ page }) => {
    // 'e' matches all three headers (name, email, member_since) before the data
    await input(page).fill('e');
    await page.waitForTimeout(DEBOUNCE);
    await input(page).press('Enter');
    const currentTh = page.locator('.subwindow table thead th.cell-find-current');
    await expect(currentTh).toBeFocused();
    await expect(currentTh).toContainText('name');
    await page.keyboard.press('Tab');
    await expect(currentTh).toContainText('email');
    await page.keyboard.press('Tab');
    await expect(currentTh).toContainText('member_since');
    // Next step crosses into the data cells
    await page.keyboard.press('Tab');
    const currentTd = page.locator('.subwindow table tbody td.cell-find-current');
    await expect(currentTd).toBeFocused();
    await expect(currentTd).toHaveText('Alice Johnson');
    // Shift+Tab goes back up to the last header match
    await page.keyboard.press('Shift+Tab');
    await expect(currentTh).toBeFocused();
    await expect(currentTh).toContainText('member_since');
  });

  test('navigating to a far off-screen match scrolls (smoothly) and focuses it', async ({ page }) => {
    await executeSQL(page,
      "WITH RECURSIVE seq(n) AS (SELECT 1 UNION ALL SELECT n + 1 FROM seq WHERE n < 300) " +
      "SELECT n, 'row' || n AS label FROM seq");
    await waitForWindow(page, '_query');
    await page.evaluate(() => app.layoutTileH());
    const bigWin = page.locator('.subwindow', { has: page.locator('.win-title', { hasText: '_query' }) });
    await bigWin.locator('.filter-input').click();
    await page.keyboard.type('row299');
    await page.waitForTimeout(DEBOUNCE);
    await page.keyboard.press('Enter');
    // The match near the bottom of 300 rows is scrolled to and focused
    // (the smooth scroll animates, so poll until it lands)
    await expect(page.locator('td.data-cell:focus')).toHaveText('row299');
    const scrollTop = await page.evaluate(() =>
      app._test.windows.find(w => w.tableName && w.tableName.startsWith('_query'))._container.scrollTop);
    expect(scrollTop).toBeGreaterThan(1000);
    // Tab wraps back to the same (only) match without breaking navigation
    await page.keyboard.press('Tab');
    await expect(page.locator('.toast', { hasText: 'wrapped' })).toBeVisible();
    await expect(page.locator('td.data-cell:focus')).toHaveText('row299');
  });

  test('no SQL autocomplete in search mode', async ({ page }) => {
    await input(page).pressSequentially('sel');
    await page.waitForTimeout(200);
    await expect(page.locator('#sql-autocomplete')).toBeHidden();
  });

  test('Enter focuses the match even when a link plugin rebuilds the source table', async ({ page }) => {
    // A second window plus a link plugin: selecting a cell in sample1 fires
    // applyLinkFilters, which rebuilds the source window's DOM (linking badges)
    // and used to silently drop the focus that Enter had just set.
    await executeSQL(page, 'SELECT * FROM sample1');
    await page.evaluate(() => {
      const cfg = {
        name: 'LinkTest',
        tables: [{ table: '.*', columns: [{ match: '^email$', display: 'upper(value)' }] }],
        links: [{ source: { table: '^sample1$', column: '^name$' }, target: { table: '.*', column: '^name$' } }],
      };
      const errors = app._test.validatePlugin(cfg);
      if (errors.length) throw new Error(errors.join(', '));
      cfg._compiled = app._test.compilePlugin(cfg);
      app._test.plugins.push(cfg);
      app._test.rebuildTransformCache();
      app._test.rebuildLinkCache();
      app._test.windows.forEach(w => app._test.rebuildTable(w));
      app.layoutTileH();
    });
    const sourceWin = page.locator('.subwindow').first();
    await sourceWin.locator('.filter-input').click();
    await page.keyboard.type('davi');
    await page.waitForTimeout(DEBOUNCE);
    await page.keyboard.press('Enter');
    await expect(page.locator('td.data-cell:focus')).toHaveText('David Brown');
    // Tab keeps navigating matches (email is plugin-formatted to uppercase)
    await page.keyboard.press('Tab');
    await expect(page.locator('td.data-cell:focus')).toHaveText('DAVID@EXAMPLE.COM');
    // A plain cell click keeps focus too, despite the link-filter rebuilds
    await sourceWin.locator('td.data-cell').first().click();
    await expect(page.locator('td.data-cell:focus')).toHaveCount(1);
  });

  test('opening Find & Replace clears an active quick search', async ({ page }) => {
    await input(page).fill('alice');
    await page.waitForTimeout(DEBOUNCE);
    await expect(page.locator('.subwindow table tbody td.cell-find-match')).toHaveCount(2);
    await page.locator('.subwindow table tbody td.data-cell').first().click();
    await page.keyboard.press('Control+f');
    await expect(page.locator('.find-input')).toBeVisible();
    await expect(input(page)).toHaveValue('');
    await expect(page.locator('.subwindow table tbody td.cell-find-match')).toHaveCount(0);
  });
});
