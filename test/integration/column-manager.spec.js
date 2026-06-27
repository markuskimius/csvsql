const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, getTableData, executeSQL } = require('../helpers');

async function getColumns(page, tableName) {
  const t = await getTableData(page, tableName);
  return t.columns;
}

async function openColManager(page) {
  await page.keyboard.press('Control+Shift+M');
  await page.waitForTimeout(300);
}

async function getColManagerItems(page) {
  return page.locator('.col-manager-item').all();
}

async function getColManagerNames(page) {
  return page.evaluate(() => {
    const items = document.querySelectorAll('.col-manager-item .col-manager-name');
    return [...items].map(el => el.textContent);
  });
}

test.describe('Column Manager', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  // ---------------------------------------------------------------------------
  // Opening and Closing
  // ---------------------------------------------------------------------------

  test.describe('Opening and Closing', () => {
    test('opens via Ctrl+Shift+M keyboard shortcut', async ({ page }) => {
      await openColManager(page);

      const managerWin = await page.evaluate(() => {
        return app._test._activeColManagerWin ? {
          tableName: app._test._activeColManagerWin._colManagerTableName
        } : null;
      });
      expect(managerWin).not.toBeNull();
      expect(managerWin.tableName).toBe('sample1');

      // Check the window title
      const title = page.locator('.subwindow.dialog .win-title');
      await expect(title).toContainText('sample1');
    });

    test('opens via context menu on column header', async ({ page }) => {
      const th = page.locator('.subwindow table thead th[data-col-idx]').first();
      await th.click({ button: 'right' });

      const menu = page.locator('.context-menu');
      await expect(menu).toBeVisible();
      await menu.locator('button', { hasText: 'Manage Columns' }).click();
      await page.waitForTimeout(300);

      const active = await page.evaluate(() => !!app._test._activeColManagerWin);
      expect(active).toBe(true);
    });

    test('opens via Edit menu button', async ({ page }) => {
      // Open Edit menu
      await page.locator('#menu-edit .menu-label').click();
      await page.waitForTimeout(200);

      // Click Manage Columns button
      await page.locator('#btn-col-manager').click();
      await page.waitForTimeout(300);

      const active = await page.evaluate(() => !!app._test._activeColManagerWin);
      expect(active).toBe(true);
    });

    test('only one column manager at a time (singleton)', async ({ page }) => {
      await openColManager(page);

      const firstId = await page.evaluate(() => app._test._activeColManagerWin.id);

      // Try opening again for the same table
      await page.keyboard.press('Control+Shift+M');
      await page.waitForTimeout(200);

      // Should focus the existing one, not create a new one
      const secondId = await page.evaluate(() => app._test._activeColManagerWin.id);
      expect(secondId).toBe(firstId);

      // Count dialog windows -- should be exactly 1
      const dialogCount = await page.locator('.subwindow.dialog').count();
      expect(dialogCount).toBe(1);
    });

    test('close via Esc key', async ({ page }) => {
      await openColManager(page);
      const active = await page.evaluate(() => !!app._test._activeColManagerWin);
      expect(active).toBe(true);

      await page.keyboard.press('Escape');
      await page.waitForTimeout(200);

      const closed = await page.evaluate(() => app._test._activeColManagerWin);
      expect(closed).toBeNull();
      await expect(page.locator('.subwindow.dialog')).toHaveCount(0);
    });

    test('close via X button', async ({ page }) => {
      await openColManager(page);

      const closeBtn = page.locator('.subwindow.dialog .btn-close');
      await closeBtn.click();
      await page.waitForTimeout(200);

      const closed = await page.evaluate(() => app._test._activeColManagerWin);
      expect(closed).toBeNull();
    });

    test('close via Ctrl+W', async ({ page }) => {
      await openColManager(page);

      // Focus the column manager window so it is active
      await page.evaluate(() => {
        app._test.focusWindow(app._test._activeColManagerWin.id);
      });
      await page.waitForTimeout(100);

      await page.keyboard.press('Control+w');
      await page.waitForTimeout(200);

      const closed = await page.evaluate(() => app._test._activeColManagerWin);
      expect(closed).toBeNull();
    });

    test('reopens after being closed', async ({ page }) => {
      await openColManager(page);
      await page.keyboard.press('Escape');
      await page.waitForTimeout(200);

      // Verify closed
      const closed = await page.evaluate(() => app._test._activeColManagerWin);
      expect(closed).toBeNull();

      // Reopen
      await openColManager(page);

      const reopened = await page.evaluate(() => !!app._test._activeColManagerWin);
      expect(reopened).toBe(true);
      await expect(page.locator('.subwindow.dialog')).toHaveCount(1);
    });

    test('closing target table closes column manager', async ({ page }) => {
      await openColManager(page);

      const active = await page.evaluate(() => !!app._test._activeColManagerWin);
      expect(active).toBe(true);

      // Close the data window (sample1) by clicking its close button.
      // The sample1 window is not the dialog -- find the non-dialog subwindow.
      const sample1Close = page.locator('.subwindow:not(.dialog) .btn-close');
      await sample1Close.click();
      await page.waitForTimeout(300);

      const closed = await page.evaluate(() => app._test._activeColManagerWin);
      expect(closed).toBeNull();
    });
  });

  // ---------------------------------------------------------------------------
  // Dialog Appearance
  // ---------------------------------------------------------------------------

  test.describe('Dialog Appearance', () => {
    test('has dialog class on subwindow', async ({ page }) => {
      await openColManager(page);
      await expect(page.locator('.subwindow.dialog')).toHaveCount(1);
    });

    test('no minimize button', async ({ page }) => {
      await openColManager(page);
      const minBtn = page.locator('.subwindow.dialog .btn-min');
      await expect(minBtn).toHaveCount(0);
    });

    test('no maximize button', async ({ page }) => {
      await openColManager(page);
      const maxBtn = page.locator('.subwindow.dialog .btn-max');
      await expect(maxBtn).toHaveCount(0);
    });

    test('has resize handles (dialogs are resizable)', async ({ page }) => {
      await openColManager(page);
      const handles = page.locator('.subwindow.dialog .resize-handle');
      await expect(handles).toHaveCount(8);
    });

    test('no status bar', async ({ page }) => {
      await openColManager(page);
      const statusbar = page.locator('.subwindow.dialog .win-statusbar');
      await expect(statusbar).toHaveCount(0);
    });

    test('close button present', async ({ page }) => {
      await openColManager(page);
      const closeBtn = page.locator('.subwindow.dialog .btn-close');
      await expect(closeBtn).toHaveCount(1);
      await expect(closeBtn).toBeVisible();
    });
  });

  // ---------------------------------------------------------------------------
  // Column List
  // ---------------------------------------------------------------------------

  test.describe('Column List', () => {
    test('shows all columns from the table', async ({ page }) => {
      await openColManager(page);

      const names = await getColManagerNames(page);
      expect(names).toEqual(['name', 'email', 'member_since']);
    });

    test('columns appear in correct order with correct index numbers', async ({ page }) => {
      await openColManager(page);

      const items = await page.evaluate(() => {
        const els = document.querySelectorAll('.col-manager-item');
        return [...els].map(el => ({
          idx: el.querySelector('.col-manager-idx').textContent,
          name: el.querySelector('.col-manager-name').textContent,
          dataIdx: el.dataset.colIdx
        }));
      });

      expect(items).toEqual([
        { idx: '1', name: 'name', dataIdx: '0' },
        { idx: '2', name: 'email', dataIdx: '1' },
        { idx: '3', name: 'member_since', dataIdx: '2' }
      ]);
    });

    test('search filters columns by name', async ({ page }) => {
      await openColManager(page);

      const searchInput = page.locator('.col-manager-search');
      await searchInput.fill('email');
      await page.waitForTimeout(200);

      const names = await getColManagerNames(page);
      expect(names).toEqual(['email']);
    });

    test('search is case-insensitive', async ({ page }) => {
      await openColManager(page);

      const searchInput = page.locator('.col-manager-search');
      await searchInput.fill('EMAIL');
      await page.waitForTimeout(200);

      const names = await getColManagerNames(page);
      expect(names).toEqual(['email']);
    });

    test('clearing search shows all columns again', async ({ page }) => {
      await openColManager(page);

      const searchInput = page.locator('.col-manager-search');
      await searchInput.fill('email');
      await page.waitForTimeout(200);

      let names = await getColManagerNames(page);
      expect(names).toEqual(['email']);

      // Clear the search
      await searchInput.fill('');
      await page.waitForTimeout(200);

      names = await getColManagerNames(page);
      expect(names).toEqual(['name', 'email', 'member_since']);
    });

    test('search with no matches shows empty list', async ({ page }) => {
      await openColManager(page);

      const searchInput = page.locator('.col-manager-search');
      await searchInput.fill('nonexistent');
      await page.waitForTimeout(200);

      const items = await page.locator('.col-manager-item').count();
      expect(items).toBe(0);
    });

    test('search matches partial column names', async ({ page }) => {
      await openColManager(page);

      const searchInput = page.locator('.col-manager-search');
      await searchInput.fill('me');
      await page.waitForTimeout(200);

      const names = await getColManagerNames(page);
      // 'name' contains 'me', 'member_since' contains 'me'
      expect(names).toEqual(['name', 'member_since']);
    });
  });

  // ---------------------------------------------------------------------------
  // Selection
  // ---------------------------------------------------------------------------

  test.describe('Selection', () => {
    test('click selects a single column (adds .selected class)', async ({ page }) => {
      await openColManager(page);

      const firstItem = page.locator('.col-manager-item').first();
      await firstItem.click();
      await page.waitForTimeout(100);

      await expect(firstItem).toHaveClass(/selected/);
    });

    test('click another column deselects previous', async ({ page }) => {
      await openColManager(page);

      const firstItem = page.locator('.col-manager-item').first();
      const secondItem = page.locator('.col-manager-item').nth(1);

      await firstItem.click();
      await page.waitForTimeout(100);
      await expect(firstItem).toHaveClass(/selected/);

      await secondItem.click();
      await page.waitForTimeout(100);
      await expect(secondItem).toHaveClass(/selected/);
      await expect(firstItem).not.toHaveClass(/selected/);
    });

    test('Ctrl+click toggles selection', async ({ page }) => {
      await openColManager(page);

      const firstItem = page.locator('.col-manager-item').first();
      const secondItem = page.locator('.col-manager-item').nth(1);
      const thirdItem = page.locator('.col-manager-item').nth(2);

      await firstItem.click();
      await page.waitForTimeout(100);

      // Meta+click (Cmd on macOS) second item -- both should be selected
      // The app checks e.ctrlKey || e.metaKey so Meta works cross-platform
      await secondItem.click({ modifiers: ['Meta'] });
      await page.waitForTimeout(100);
      await expect(firstItem).toHaveClass(/selected/);
      await expect(secondItem).toHaveClass(/selected/);

      // Meta+click third item -- all three should be selected
      await thirdItem.click({ modifiers: ['Meta'] });
      await page.waitForTimeout(100);
      await expect(firstItem).toHaveClass(/selected/);
      await expect(secondItem).toHaveClass(/selected/);
      await expect(thirdItem).toHaveClass(/selected/);
    });

    test('Shift+click selects range', async ({ page }) => {
      await openColManager(page);

      const firstItem = page.locator('.col-manager-item').first();
      const thirdItem = page.locator('.col-manager-item').nth(2);

      await firstItem.click();
      await page.waitForTimeout(100);

      await thirdItem.click({ modifiers: ['Shift'] });
      await page.waitForTimeout(100);

      // All three items should be selected
      const items = page.locator('.col-manager-item');
      await expect(items.nth(0)).toHaveClass(/selected/);
      await expect(items.nth(1)).toHaveClass(/selected/);
      await expect(items.nth(2)).toHaveClass(/selected/);
    });

    test('Arrow keys navigate selection', async ({ page }) => {
      await openColManager(page);

      // Focus the list and select first item
      const firstItem = page.locator('.col-manager-item').first();
      await firstItem.click();
      await page.waitForTimeout(100);

      // Focus the list element for keyboard events
      await page.locator('.col-manager-list').focus();

      // Arrow down should select second item
      await page.keyboard.press('ArrowDown');
      await page.waitForTimeout(100);

      const items = page.locator('.col-manager-item');
      await expect(items.nth(0)).not.toHaveClass(/selected/);
      await expect(items.nth(1)).toHaveClass(/selected/);

      // Arrow down again should select third item
      await page.keyboard.press('ArrowDown');
      await page.waitForTimeout(100);
      await expect(items.nth(1)).not.toHaveClass(/selected/);
      await expect(items.nth(2)).toHaveClass(/selected/);

      // Arrow up should go back to second
      await page.keyboard.press('ArrowUp');
      await page.waitForTimeout(100);
      await expect(items.nth(2)).not.toHaveClass(/selected/);
      await expect(items.nth(1)).toHaveClass(/selected/);
    });

    test('Shift+Arrow extends selection', async ({ page }) => {
      await openColManager(page);

      const firstItem = page.locator('.col-manager-item').first();
      await firstItem.click();
      await page.waitForTimeout(100);

      await page.locator('.col-manager-list').focus();

      // Shift+ArrowDown extends selection to include second item
      await page.keyboard.press('Shift+ArrowDown');
      await page.waitForTimeout(100);

      const items = page.locator('.col-manager-item');
      await expect(items.nth(0)).toHaveClass(/selected/);
      await expect(items.nth(1)).toHaveClass(/selected/);
    });
  });

  // ---------------------------------------------------------------------------
  // Reorder via Alt+Up/Down
  // ---------------------------------------------------------------------------

  test.describe('Reorder via Alt+Up/Down', () => {
    test('Alt+Up moves selected column up', async ({ page }) => {
      await openColManager(page);

      // Select the second item (email)
      const secondItem = page.locator('.col-manager-item').nth(1);
      await secondItem.click();
      await page.waitForTimeout(100);

      await page.locator('.col-manager-list').focus();
      await page.keyboard.press('Alt+ArrowUp');
      await page.waitForTimeout(300);

      // Table columns should now be email, name, member_since
      const cols = await getColumns(page, 'sample1');
      expect(cols).toEqual(['email', 'name', 'member_since']);

      // Column manager list should also reflect the new order
      const names = await getColManagerNames(page);
      expect(names).toEqual(['email', 'name', 'member_since']);
    });

    test('Alt+Down moves selected column down', async ({ page }) => {
      await openColManager(page);

      // Select the first item (name)
      const firstItem = page.locator('.col-manager-item').first();
      await firstItem.click();
      await page.waitForTimeout(100);

      await page.locator('.col-manager-list').focus();
      await page.keyboard.press('Alt+ArrowDown');
      await page.waitForTimeout(300);

      const cols = await getColumns(page, 'sample1');
      expect(cols).toEqual(['email', 'name', 'member_since']);

      const names = await getColManagerNames(page);
      expect(names).toEqual(['email', 'name', 'member_since']);
    });

    test('boundary check: cannot move past top', async ({ page }) => {
      await openColManager(page);

      // Select the first item (name)
      const firstItem = page.locator('.col-manager-item').first();
      await firstItem.click();
      await page.waitForTimeout(100);

      await page.locator('.col-manager-list').focus();
      await page.keyboard.press('Alt+ArrowUp');
      await page.waitForTimeout(200);

      // Order should be unchanged
      const cols = await getColumns(page, 'sample1');
      expect(cols).toEqual(['name', 'email', 'member_since']);
    });

    test('boundary check: cannot move past bottom', async ({ page }) => {
      await openColManager(page);

      // Select the last item (member_since)
      const lastItem = page.locator('.col-manager-item').last();
      await lastItem.click();
      await page.waitForTimeout(100);

      await page.locator('.col-manager-list').focus();
      await page.keyboard.press('Alt+ArrowDown');
      await page.waitForTimeout(200);

      // Order should be unchanged
      const cols = await getColumns(page, 'sample1');
      expect(cols).toEqual(['name', 'email', 'member_since']);
    });

    test('table columns update in real time', async ({ page }) => {
      await openColManager(page);

      // Select email (index 1) and move it down twice
      const secondItem = page.locator('.col-manager-item').nth(1);
      await secondItem.click();
      await page.waitForTimeout(100);

      await page.locator('.col-manager-list').focus();
      await page.keyboard.press('Alt+ArrowDown');
      await page.waitForTimeout(200);

      // After first move: name, member_since, email
      let cols = await getColumns(page, 'sample1');
      expect(cols).toEqual(['name', 'member_since', 'email']);

      // The table headers in the DOM should also reflect this
      const headers = await page.evaluate(() => {
        const ths = document.querySelectorAll('.subwindow:not(.dialog) table thead th[data-col-idx]');
        return [...ths].map(th => th.textContent.replace(/[≡☰]/g, '').trim());
      });
      expect(headers).toEqual(['name', 'member_since', 'email']);
    });
  });

  // ---------------------------------------------------------------------------
  // Double-click to Scroll
  // ---------------------------------------------------------------------------

  test.describe('Double-click to scroll', () => {
    test('double-clicking a column name scrolls the table to show that column', async ({ page }) => {
      await openColManager(page);

      // Double-click on member_since (last column)
      const lastItem = page.locator('.col-manager-item').last();
      await lastItem.dblclick();
      await page.waitForTimeout(300);

      // After double-click, the target window's selectedCol should be set
      const selectedCol = await page.evaluate(() => {
        const win = app._test.windows.find(w => w.tableName === 'sample1' && !w._colManagerTableName);
        return win ? win.selectedCol : null;
      });
      expect(selectedCol).toBe('member_since');
    });
  });

  // ---------------------------------------------------------------------------
  // Dimming
  // ---------------------------------------------------------------------------

  test.describe('Dimming', () => {
    test('other windows get dimmed class when column manager opens', async ({ page }) => {
      // Open a second table via SQL query so we have another window
      await executeSQL(page, "SELECT 'a' AS x");
      await page.waitForTimeout(300);

      // Focus sample1 and open col manager
      await page.evaluate(() => {
        const win = app._test.windows.find(w => w.tableName === 'sample1');
        app._test.focusWindow(win.id);
      });
      await page.waitForTimeout(200);

      await openColManager(page);

      // The query result window should be dimmed
      const dimmedCount = await page.locator('.col-manager-dimmed').count();
      expect(dimmedCount).toBeGreaterThan(0);
    });

    test('target table window is NOT dimmed', async ({ page }) => {
      // Open a second window
      await executeSQL(page, "SELECT 'a' AS x");
      await page.waitForTimeout(300);

      // Focus sample1 and open col manager for it
      await page.evaluate(() => {
        const win = app._test.windows.find(w => w.tableName === 'sample1');
        app._test.focusWindow(win.id);
      });
      await page.waitForTimeout(200);

      await openColManager(page);

      // Check that sample1 window does not have the dimmed class
      const targetDimmed = await page.evaluate(() => {
        const win = app._test.windows.find(w => w.tableName === 'sample1' && !w._colManagerTableName);
        return win ? win.el.classList.contains('col-manager-dimmed') : null;
      });
      expect(targetDimmed).toBe(false);
    });

    test('dimming clears when column manager closes', async ({ page }) => {
      // Open a second window
      await executeSQL(page, "SELECT 'a' AS x");
      await page.waitForTimeout(300);

      await page.evaluate(() => {
        const win = app._test.windows.find(w => w.tableName === 'sample1');
        app._test.focusWindow(win.id);
      });
      await page.waitForTimeout(200);

      await openColManager(page);

      // Verify dimming is active
      let dimmedCount = await page.locator('.col-manager-dimmed').count();
      expect(dimmedCount).toBeGreaterThan(0);

      // Close the column manager
      await page.keyboard.press('Escape');
      await page.waitForTimeout(200);

      // Dimming should be cleared
      dimmedCount = await page.locator('.col-manager-dimmed').count();
      expect(dimmedCount).toBe(0);
    });

    test('dimmed windows are darkened (filter), not translucent (opacity stays 1)', async ({ page }) => {
      // Open a second window so there's something to dim
      await executeSQL(page, "SELECT 'a' AS x");
      await page.waitForTimeout(300);

      await page.evaluate(() => {
        const win = app._test.windows.find(w => w.tableName === 'sample1');
        app._test.focusWindow(win.id);
      });
      await page.waitForTimeout(200);

      await openColManager(page);

      const style = await page.evaluate(() => {
        const el = document.querySelector('.col-manager-dimmed');
        if (!el) return null;
        const s = getComputedStyle(el);
        return { opacity: s.opacity, filter: s.filter };
      });
      expect(style).not.toBeNull();
      // The fix: dim via a filter, NOT opacity — so the window stays opaque.
      expect(style.opacity).toBe('1');
      expect(style.filter).toContain('brightness');
      expect(style.filter).toContain('grayscale');
    });
  });

  // ---------------------------------------------------------------------------
  // Undo Integration
  // ---------------------------------------------------------------------------

  test.describe('Undo Integration', () => {
    test('reorders via column manager batch into single undo entry on close', async ({ page }) => {
      await openColManager(page);

      // Make multiple reorders: move email up, then move member_since up
      const secondItem = page.locator('.col-manager-item').nth(1);
      await secondItem.click();
      await page.waitForTimeout(100);

      await page.locator('.col-manager-list').focus();
      await page.keyboard.press('Alt+ArrowUp');
      await page.waitForTimeout(200);

      // Now select what was member_since (now at index 2) and move up
      const thirdItem = page.locator('.col-manager-item').nth(2);
      await thirdItem.click();
      await page.waitForTimeout(100);

      await page.locator('.col-manager-list').focus();
      await page.keyboard.press('Alt+ArrowUp');
      await page.waitForTimeout(200);

      // Columns should be: email, member_since, name
      let cols = await getColumns(page, 'sample1');
      expect(cols).toEqual(['email', 'member_since', 'name']);

      // Close the column manager
      await page.keyboard.press('Escape');
      await page.waitForTimeout(200);

      // Check that only a single undo entry was added
      const undoStackSize = await page.evaluate(() => {
        const t = app._test.tables['sample1'];
        return t._undoStack ? t._undoStack.length : 0;
      });
      expect(undoStackSize).toBe(1);

      // The undo entry should be a reorderColumn type
      const undoType = await page.evaluate(() => {
        const t = app._test.tables['sample1'];
        return t._undoStack[t._undoStack.length - 1].type;
      });
      expect(undoType).toBe('reorderColumn');
    });

    test('Ctrl+Z after closing reverts all reorders at once', async ({ page }) => {
      await openColManager(page);

      // Move email up
      const secondItem = page.locator('.col-manager-item').nth(1);
      await secondItem.click();
      await page.waitForTimeout(100);

      await page.locator('.col-manager-list').focus();
      await page.keyboard.press('Alt+ArrowUp');
      await page.waitForTimeout(200);

      // Move member_since up
      const thirdItem = page.locator('.col-manager-item').nth(2);
      await thirdItem.click();
      await page.waitForTimeout(100);

      await page.locator('.col-manager-list').focus();
      await page.keyboard.press('Alt+ArrowUp');
      await page.waitForTimeout(200);

      // Close the column manager
      await page.keyboard.press('Escape');
      await page.waitForTimeout(200);

      let cols = await getColumns(page, 'sample1');
      expect(cols).toEqual(['email', 'member_since', 'name']);

      // Focus the data window and undo
      await page.evaluate(() => {
        const win = app._test.windows.find(w => w.tableName === 'sample1');
        app._test.focusWindow(win.id);
      });
      await page.waitForTimeout(100);

      await page.keyboard.press('Control+z');
      await page.waitForTimeout(300);

      // All reorders should be reverted in a single undo
      cols = await getColumns(page, 'sample1');
      expect(cols).toEqual(['name', 'email', 'member_since']);
    });

    test('no undo entry if no changes were made', async ({ page }) => {
      // Record initial undo stack size
      const initialSize = await page.evaluate(() => {
        const t = app._test.tables['sample1'];
        return t._undoStack ? t._undoStack.length : 0;
      });

      await openColManager(page);

      // Close without making any changes
      await page.keyboard.press('Escape');
      await page.waitForTimeout(200);

      const afterSize = await page.evaluate(() => {
        const t = app._test.tables['sample1'];
        return t._undoStack ? t._undoStack.length : 0;
      });
      expect(afterSize).toBe(initialSize);
    });
  });
});
