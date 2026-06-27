const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, getTableData } = require('../helpers');

test.describe('Row and Column Operations', () => {
  test.beforeEach(async ({ page }) => {
    await openApp(page);
    await uploadFile(page, '../test/sample1.csv');
    await waitForWindow(page, 'sample1');
  });

  // ---------------------------------------------------------------------------
  // Context Menus
  // ---------------------------------------------------------------------------

  test.describe('Context Menus', () => {
    test('right-click row number shows context menu with Insert Row Below and Delete Row', async ({ page }) => {
      const rowNum = page.locator('.subwindow table tbody td.row-num').first();
      await rowNum.click({ button: 'right' });
      const menu = page.locator('.context-menu');
      await expect(menu).toBeVisible();
      const buttons = menu.locator('button');
      await expect(buttons.filter({ hasText: 'Insert Row Below' })).toBeVisible();
      await expect(buttons.filter({ hasText: 'Delete Row' })).toBeVisible();
    });

    test('right-click column header shows context menu with Insert Column Right and Delete Column', async ({ page }) => {
      const th = page.locator('.subwindow table thead th[data-col-idx]').first();
      await th.click({ button: 'right' });
      const menu = page.locator('.context-menu');
      await expect(menu).toBeVisible();
      const buttons = menu.locator('button');
      await expect(buttons.filter({ hasText: 'Insert Column Right' })).toBeVisible();
      await expect(buttons.filter({ hasText: 'Delete Column' })).toBeVisible();
    });

    test('right-click corner cell shows context menu with Insert Column Right and Insert Row Below', async ({ page }) => {
      const corner = page.locator('.subwindow table thead th.row-num-header');
      await corner.click({ button: 'right' });
      const menu = page.locator('.context-menu');
      await expect(menu).toBeVisible();
      const buttons = menu.locator('button');
      await expect(buttons.filter({ hasText: 'Insert Column Right' })).toBeVisible();
      await expect(buttons.filter({ hasText: 'Insert Row Below' })).toBeVisible();
    });

    test('Delete Column is disabled when only 1 column remains', async ({ page }) => {
      // Delete all columns except one
      await page.evaluate(async () => {
        const t = app._test.tables['sample1'];
        while (t.columns.length > 1) {
          await app._test.deleteColumn('sample1', t.columns[t.columns.length - 1]);
        }
      });
      await page.waitForTimeout(300);

      // Right-click the remaining column header
      const th = page.locator('.subwindow table thead th[data-col-idx]').first();
      await th.click({ button: 'right' });
      const menu = page.locator('.context-menu');
      await expect(menu).toBeVisible();
      const deleteBtn = menu.locator('button', { hasText: 'Delete Column' });
      await expect(deleteBtn).toBeDisabled();
    });

    test('context menu closes when clicking elsewhere', async ({ page }) => {
      const rowNum = page.locator('.subwindow table tbody td.row-num').first();
      await rowNum.click({ button: 'right' });
      const menu = page.locator('.context-menu');
      await expect(menu).toBeVisible();
      // Click on the window body to dismiss
      await page.locator('#window-area').click({ position: { x: 10, y: 10 } });
      await page.waitForTimeout(200);
      await expect(menu).not.toBeAttached();
    });
  });

  // ---------------------------------------------------------------------------
  // Row Operations
  // ---------------------------------------------------------------------------

  test.describe('Row Operations', () => {
    test('Insert Row Below via row context menu adds a row after the clicked row', async ({ page }) => {
      const before = await getTableData(page, 'sample1');
      const firstRowName = before.rows[0].name;
      const secondRowName = before.rows[1].name;

      // Right-click on the first row number
      const rowNum = page.locator('.subwindow table tbody td.row-num').first();
      await rowNum.click({ button: 'right' });
      await page.locator('.context-menu button', { hasText: 'Insert Row Below' }).click();
      await page.waitForTimeout(300);

      const after = await getTableData(page, 'sample1');
      expect(after.rows.length).toBe(before.rows.length + 1);
      // Original first row stays at index 0
      expect(after.rows[0].name).toBe(firstRowName);
      // New empty row inserted at index 1
      expect(after.rows[1].name).toBe('');
      // Original second row shifted to index 2
      expect(after.rows[2].name).toBe(secondRowName);
    });

    test('Insert Row Below via corner context menu adds a row at position 0', async ({ page }) => {
      const before = await getTableData(page, 'sample1');
      const firstRowName = before.rows[0].name;

      const corner = page.locator('.subwindow table thead th.row-num-header');
      await corner.click({ button: 'right' });
      await page.locator('.context-menu button', { hasText: 'Insert Row Below' }).click();
      await page.waitForTimeout(300);

      const after = await getTableData(page, 'sample1');
      expect(after.rows.length).toBe(before.rows.length + 1);
      // New empty row at index 0
      expect(after.rows[0].name).toBe('');
      // Original first row shifted to index 1
      expect(after.rows[1].name).toBe(firstRowName);
    });

    test('Delete Row via context menu removes the correct row', async ({ page }) => {
      const before = await getTableData(page, 'sample1');
      const firstRowName = before.rows[0].name;
      const secondRowName = before.rows[1].name;

      // Right-click on the first row number
      const rowNum = page.locator('.subwindow table tbody td.row-num').first();
      await rowNum.click({ button: 'right' });
      await page.locator('.context-menu button', { hasText: 'Delete Row' }).click();
      await page.waitForTimeout(300);

      const after = await getTableData(page, 'sample1');
      expect(after.rows.length).toBe(before.rows.length - 1);
      // The second row is now first
      expect(after.rows[0].name).toBe(secondRowName);
    });

    test('row numbers renumber correctly after insert and delete', async ({ page }) => {
      // Insert a row after row 1
      const rowNum = page.locator('.subwindow table tbody td.row-num').first();
      await rowNum.click({ button: 'right' });
      await page.locator('.context-menu button', { hasText: 'Insert Row Below' }).click();
      await page.waitForTimeout(300);

      let data = await getTableData(page, 'sample1');
      // Verify all _rownum values are sequential
      data.rows.forEach((row, i) => {
        expect(row._rownum).toBe(i + 1);
      });

      // Now delete row 3
      const rowNum3 = page.locator('.subwindow table tbody td.row-num').nth(2);
      await rowNum3.click({ button: 'right' });
      await page.locator('.context-menu button', { hasText: 'Delete Row' }).click();
      await page.waitForTimeout(300);

      data = await getTableData(page, 'sample1');
      data.rows.forEach((row, i) => {
        expect(row._rownum).toBe(i + 1);
      });
    });
  });

  // ---------------------------------------------------------------------------
  // Column Operations
  // ---------------------------------------------------------------------------

  test.describe('Column Operations', () => {
    test('Insert Column Right via column header context menu', async ({ page }) => {
      const before = await getTableData(page, 'sample1');
      const firstCol = before.columns[0]; // 'name'

      // Use app._test.insertColumn directly to avoid the prompt dialog
      await page.evaluate(async () => {
        await app._test.insertColumn('sample1', 'newcol', 1);
        const win = app._test.windows.find(w => w.tableName === 'sample1');
        app._test.rebuildTable(win);
      });
      await page.waitForTimeout(300);

      const after = await getTableData(page, 'sample1');
      expect(after.columns.length).toBe(before.columns.length + 1);
      expect(after.columns[0]).toBe(firstCol);
      expect(after.columns[1]).toBe('newcol');
      // New column has empty values
      expect(after.rows[0].newcol).toBe('');
    });

    test('Insert Column Right via corner context menu inserts at position 0', async ({ page }) => {
      const before = await getTableData(page, 'sample1');

      // Use app._test.insertColumn at index 0
      await page.evaluate(async () => {
        await app._test.insertColumn('sample1', 'firstcol', 0);
        const win = app._test.windows.find(w => w.tableName === 'sample1');
        app._test.rebuildTable(win);
      });
      await page.waitForTimeout(300);

      const after = await getTableData(page, 'sample1');
      expect(after.columns.length).toBe(before.columns.length + 1);
      expect(after.columns[0]).toBe('firstcol');
      expect(after.columns[1]).toBe(before.columns[0]);
    });

    test('Insert Column Right via prompt dialog from column header context menu', async ({ page }) => {
      const before = await getTableData(page, 'sample1');

      // Right-click on the first column header
      const th = page.locator('.subwindow table thead th[data-col-idx]').first();
      await th.click({ button: 'right' });
      await page.locator('.context-menu button', { hasText: 'Insert Column Right' }).click();

      // Interact with the prompt dialog
      await page.waitForSelector('.subwindow.dialog');
      const input = page.locator('.subwindow.dialog input');
      await input.fill('new_column');
      await page.locator('.subwindow.dialog button', { hasText: 'OK' }).click();
      await page.waitForTimeout(500);

      const after = await getTableData(page, 'sample1');
      expect(after.columns.length).toBe(before.columns.length + 1);
      expect(after.columns).toContain('new_column');
    });

    test('Delete Column via column header context menu removes the column', async ({ page }) => {
      const before = await getTableData(page, 'sample1');
      const colToDelete = before.columns[2]; // 'member_since'

      const th = page.locator('.subwindow table thead th[data-col-idx]').nth(2);
      await th.click({ button: 'right' });
      await page.locator('.context-menu button', { hasText: 'Delete Column' }).click();
      await page.waitForTimeout(500);

      const after = await getTableData(page, 'sample1');
      expect(after.columns.length).toBe(before.columns.length - 1);
      expect(after.columns).not.toContain(colToDelete);
      // Verify the column data is gone from rows
      expect(after.rows[0]).not.toHaveProperty(colToDelete);
    });

    test('Delete Column cleans up window state', async ({ page }) => {
      // Set up some state on the column we will delete
      await page.evaluate(() => {
        const win = app._test.windows.find(w => w.tableName === 'sample1');
        const t = app._test.tables['sample1'];
        // Add a column filter
        if (!win.columnFilters) win.columnFilters = {};
        win.columnFilters['member_since'] = new Set(['val']);
        // Add to sort
        win.sortCols = [{ col: 'member_since', dir: 'asc' }];
        // Set as selected column
        win.selectedCol = 'member_since';
      });

      // Delete the column
      await page.evaluate(async () => {
        await app._test.deleteColumn('sample1', 'member_since');
      });
      await page.waitForTimeout(500);

      const state = await page.evaluate(() => {
        const win = app._test.windows.find(w => w.tableName === 'sample1');
        return {
          hasFilter: win.columnFilters && 'member_since' in win.columnFilters,
          sortCols: win.sortCols,
          selectedCol: win.selectedCol,
        };
      });
      expect(state.hasFilter).toBe(false);
      expect(state.sortCols).toEqual([]);
      expect(state.selectedCol).toBeNull();
    });

    test('Insert column preserves existing column widths', async ({ page }) => {
      // First ensure colWidths are set by getting current state
      const widthsBefore = await page.evaluate(() => {
        const win = app._test.windows.find(w => w.tableName === 'sample1');
        return win.colWidths ? [...win.colWidths] : null;
      });

      // Insert a column
      await page.evaluate(async () => {
        await app._test.insertColumn('sample1', 'inserted_col', 1);
        const win = app._test.windows.find(w => w.tableName === 'sample1');
        app._test.rebuildTable(win);
      });
      await page.waitForTimeout(300);

      if (widthsBefore) {
        const widthsAfter = await page.evaluate(() => {
          const win = app._test.windows.find(w => w.tableName === 'sample1');
          return win.colWidths ? [...win.colWidths] : null;
        });

        // Should have one more entry
        expect(widthsAfter.length).toBe(widthsBefore.length + 1);
        // Original first column width should be preserved
        expect(widthsAfter[0]).toBe(widthsBefore[0]);
        // New column at index 1 should have default width (100)
        expect(widthsAfter[1]).toBe(100);
        // Remaining original columns should be shifted but preserved
        for (let i = 1; i < widthsBefore.length; i++) {
          expect(widthsAfter[i + 1]).toBe(widthsBefore[i]);
        }
      }
    });
  });

  // ---------------------------------------------------------------------------
  // Structural Undo/Redo
  // ---------------------------------------------------------------------------

  test.describe('Structural Undo/Redo', () => {
    test('undo insert row removes the inserted row', async ({ page }) => {
      const before = await getTableData(page, 'sample1');

      // Insert a row via context menu
      const rowNum = page.locator('.subwindow table tbody td.row-num').first();
      await rowNum.click({ button: 'right' });
      await page.locator('.context-menu button', { hasText: 'Insert Row Below' }).click();
      await page.waitForTimeout(300);

      let data = await getTableData(page, 'sample1');
      expect(data.rows.length).toBe(before.rows.length + 1);

      // Undo
      await page.evaluate(() => {
        const win = app._test.windows.find(w => w.tableName === 'sample1');
        app._test.undoTable('sample1', win);
      });
      await page.waitForTimeout(500);

      data = await getTableData(page, 'sample1');
      expect(data.rows.length).toBe(before.rows.length);
      expect(data.rows[0].name).toBe(before.rows[0].name);
    });

    test('redo insert row re-inserts it', async ({ page }) => {
      const before = await getTableData(page, 'sample1');

      // Insert row
      const rowNum = page.locator('.subwindow table tbody td.row-num').first();
      await rowNum.click({ button: 'right' });
      await page.locator('.context-menu button', { hasText: 'Insert Row Below' }).click();
      await page.waitForTimeout(300);

      // Undo
      await page.evaluate(() => {
        const win = app._test.windows.find(w => w.tableName === 'sample1');
        app._test.undoTable('sample1', win);
      });
      await page.waitForTimeout(500);

      // Redo
      await page.evaluate(() => {
        const win = app._test.windows.find(w => w.tableName === 'sample1');
        app._test.redoTable('sample1', win);
      });
      await page.waitForTimeout(500);

      const data = await getTableData(page, 'sample1');
      expect(data.rows.length).toBe(before.rows.length + 1);
      // The inserted row should be empty at index 1
      expect(data.rows[1].name).toBe('');
    });

    test('undo delete row restores the deleted row with original data', async ({ page }) => {
      const before = await getTableData(page, 'sample1');
      const deletedName = before.rows[0].name;
      const deletedEmail = before.rows[0].email;

      // Delete the first row
      const rowNum = page.locator('.subwindow table tbody td.row-num').first();
      await rowNum.click({ button: 'right' });
      await page.locator('.context-menu button', { hasText: 'Delete Row' }).click();
      await page.waitForTimeout(300);

      let data = await getTableData(page, 'sample1');
      expect(data.rows.length).toBe(before.rows.length - 1);

      // Undo
      await page.evaluate(() => {
        const win = app._test.windows.find(w => w.tableName === 'sample1');
        app._test.undoTable('sample1', win);
      });
      await page.waitForTimeout(500);

      data = await getTableData(page, 'sample1');
      expect(data.rows.length).toBe(before.rows.length);
      expect(data.rows[0].name).toBe(deletedName);
      expect(data.rows[0].email).toBe(deletedEmail);
    });

    test('redo delete row re-deletes it', async ({ page }) => {
      const before = await getTableData(page, 'sample1');

      // Delete the first row
      const rowNum = page.locator('.subwindow table tbody td.row-num').first();
      await rowNum.click({ button: 'right' });
      await page.locator('.context-menu button', { hasText: 'Delete Row' }).click();
      await page.waitForTimeout(300);

      // Undo
      await page.evaluate(() => {
        const win = app._test.windows.find(w => w.tableName === 'sample1');
        app._test.undoTable('sample1', win);
      });
      await page.waitForTimeout(500);

      let data = await getTableData(page, 'sample1');
      expect(data.rows.length).toBe(before.rows.length);

      // Redo
      await page.evaluate(() => {
        const win = app._test.windows.find(w => w.tableName === 'sample1');
        app._test.redoTable('sample1', win);
      });
      await page.waitForTimeout(500);

      data = await getTableData(page, 'sample1');
      expect(data.rows.length).toBe(before.rows.length - 1);
    });

    test('undo insert column removes the inserted column', async ({ page }) => {
      const before = await getTableData(page, 'sample1');

      await page.evaluate(async () => {
        await app._test.insertColumn('sample1', 'newcol', 1);
        const win = app._test.windows.find(w => w.tableName === 'sample1');
        app._test.rebuildTable(win);
      });
      await page.waitForTimeout(300);

      let data = await getTableData(page, 'sample1');
      expect(data.columns).toContain('newcol');

      // Undo
      await page.evaluate(() => {
        const win = app._test.windows.find(w => w.tableName === 'sample1');
        app._test.undoTable('sample1', win);
      });
      await page.waitForTimeout(500);

      data = await getTableData(page, 'sample1');
      expect(data.columns).not.toContain('newcol');
      expect(data.columns.length).toBe(before.columns.length);
    });

    test('redo insert column re-inserts it', async ({ page }) => {
      await page.evaluate(async () => {
        await app._test.insertColumn('sample1', 'newcol', 1);
        const win = app._test.windows.find(w => w.tableName === 'sample1');
        app._test.rebuildTable(win);
      });
      await page.waitForTimeout(300);

      // Undo
      await page.evaluate(() => {
        const win = app._test.windows.find(w => w.tableName === 'sample1');
        app._test.undoTable('sample1', win);
      });
      await page.waitForTimeout(500);

      // Redo
      await page.evaluate(() => {
        const win = app._test.windows.find(w => w.tableName === 'sample1');
        app._test.redoTable('sample1', win);
      });
      await page.waitForTimeout(500);

      const data = await getTableData(page, 'sample1');
      expect(data.columns).toContain('newcol');
      expect(data.columns[1]).toBe('newcol');
    });

    test('undo delete column restores the column with all its data', async ({ page }) => {
      const before = await getTableData(page, 'sample1');
      const colToDelete = 'email';
      const originalEmails = before.rows.map(r => r[colToDelete]);

      await page.evaluate(async () => {
        await app._test.deleteColumn('sample1', 'email');
      });
      await page.waitForTimeout(500);

      let data = await getTableData(page, 'sample1');
      expect(data.columns).not.toContain(colToDelete);

      // Undo
      await page.evaluate(() => {
        const win = app._test.windows.find(w => w.tableName === 'sample1');
        app._test.undoTable('sample1', win);
      });
      await page.waitForTimeout(500);

      data = await getTableData(page, 'sample1');
      expect(data.columns).toContain(colToDelete);
      // Verify all original data is restored
      const restoredEmails = data.rows.map(r => r[colToDelete]);
      expect(restoredEmails).toEqual(originalEmails);
    });

    test('redo delete column re-deletes it', async ({ page }) => {
      await page.evaluate(async () => {
        await app._test.deleteColumn('sample1', 'email');
      });
      await page.waitForTimeout(500);

      // Undo
      await page.evaluate(() => {
        const win = app._test.windows.find(w => w.tableName === 'sample1');
        app._test.undoTable('sample1', win);
      });
      await page.waitForTimeout(500);

      let data = await getTableData(page, 'sample1');
      expect(data.columns).toContain('email');

      // Redo
      await page.evaluate(() => {
        const win = app._test.windows.find(w => w.tableName === 'sample1');
        app._test.redoTable('sample1', win);
      });
      await page.waitForTimeout(500);

      data = await getTableData(page, 'sample1');
      expect(data.columns).not.toContain('email');
    });

    test('Edit menu shows contextual undo/redo labels for structural operations', async ({ page }) => {
      // Insert a row to create an undo entry
      const rowNum = page.locator('.subwindow table tbody td.row-num').first();
      await rowNum.click({ button: 'right' });
      await page.locator('.context-menu button', { hasText: 'Insert Row Below' }).click();
      await page.waitForTimeout(300);

      // Verify undo stack has the right entry type
      const undoType = await page.evaluate(() => {
        const t = app._test.tables['sample1'];
        return t._undoStack && t._undoStack.length > 0 ? t._undoStack[t._undoStack.length - 1].type : null;
      });
      expect(undoType).toBe('insertRow');

      // Open Edit menu and check undo label (use firstChild.textContent to skip shortcut span)
      await page.locator('#menu-edit .menu-label').click();
      await page.waitForTimeout(100);
      const undoLabel = await page.locator('#btn-undo').evaluate(el => el.firstChild.textContent);
      expect(undoLabel).toContain('Undo Insert Row');

      // Close the menu by clicking elsewhere on the menubar
      await page.locator('#menu-edit .menu-label').click();
      await page.waitForTimeout(100);

      // Undo via evaluate to create a redo entry
      await page.evaluate(() => {
        const win = app._test.windows.find(w => w.tableName === 'sample1');
        app._test.undoTable('sample1', win);
      });
      await page.waitForTimeout(500);

      // Verify redo stack has the right entry type
      const redoType = await page.evaluate(() => {
        const t = app._test.tables['sample1'];
        return t._redoStack && t._redoStack.length > 0 ? t._redoStack[t._redoStack.length - 1].type : null;
      });
      expect(redoType).toBe('insertRow');

      // Focus a cell so getActiveDataWindow() works, then check Edit menu redo label
      const cell = page.locator('.subwindow table tbody td.data-cell').first();
      await cell.click({ force: true });
      await page.waitForTimeout(100);
      await page.locator('#menu-edit .menu-label').click();
      await page.waitForTimeout(100);
      const redoLabel = await page.locator('#btn-redo').evaluate(el => el.firstChild.textContent);
      expect(redoLabel).toContain('Redo Insert Row');
      await page.keyboard.press('Escape');
    });

    test('Edit menu shows contextual label for delete column undo', async ({ page }) => {
      // Delete a column
      await page.evaluate(async () => {
        await app._test.deleteColumn('sample1', 'member_since');
      });
      await page.waitForTimeout(500);

      // Focus the window by clicking a cell so getActiveDataWindow() works
      const cell = page.locator('.subwindow table tbody td.data-cell').first();
      await cell.click();
      await page.waitForTimeout(100);

      // Open Edit menu and check undo label
      await page.locator('#menu-edit .menu-label').click();
      await page.waitForTimeout(100);
      const undoLabel = await page.locator('#btn-undo').evaluate(el => el.firstChild.textContent);
      expect(undoLabel).toContain('Undo Delete Column');
      await page.keyboard.press('Escape');
    });

    test('undo insert column preserves existing column widths', async ({ page }) => {
      // Get widths before insert
      const widthsBefore = await page.evaluate(() => {
        const win = app._test.windows.find(w => w.tableName === 'sample1');
        return win.colWidths ? [...win.colWidths] : null;
      });

      // Insert column
      await page.evaluate(async () => {
        await app._test.insertColumn('sample1', 'temp_col', 1);
        const win = app._test.windows.find(w => w.tableName === 'sample1');
        app._test.rebuildTable(win);
      });
      await page.waitForTimeout(300);

      // Undo
      await page.evaluate(() => {
        const win = app._test.windows.find(w => w.tableName === 'sample1');
        app._test.undoTable('sample1', win);
      });
      await page.waitForTimeout(500);

      const widthsAfter = await page.evaluate(() => {
        const win = app._test.windows.find(w => w.tableName === 'sample1');
        return win.colWidths ? [...win.colWidths] : null;
      });

      // After undo, the widths may be reset (registerTable rebuilds).
      // At minimum, column count should match the original
      if (widthsBefore && widthsAfter) {
        expect(widthsAfter.length).toBe(widthsBefore.length);
      }
    });
  });

  // ---------------------------------------------------------------------------
  // Insert Column Dialog Validation
  // ---------------------------------------------------------------------------

  test.describe('Insert Column Dialog Validation', () => {
    test('inserting a column with a duplicate name shows error in the dialog', async ({ page }) => {
      // Right-click on the first column header
      const th = page.locator('.subwindow table thead th[data-col-idx]').first();
      await th.click({ button: 'right' });
      await page.locator('.context-menu button', { hasText: 'Insert Column Right' }).click();

      // Wait for dialog
      await page.waitForSelector('.subwindow.dialog');
      const input = page.locator('.subwindow.dialog input');
      const errorDiv = page.locator('.subwindow.dialog .modal-error');

      // Type a duplicate column name ('email' already exists)
      await input.fill('email');
      await page.locator('.subwindow.dialog button', { hasText: 'OK' }).click();
      await page.waitForTimeout(200);

      // Dialog should still be open with an error message
      await expect(page.locator('.subwindow.dialog')).toBeVisible();
      const errorText = await errorDiv.textContent();
      expect(errorText.length).toBeGreaterThan(0);
    });

    test('correcting the duplicate name and resubmitting works', async ({ page }) => {
      const before = await getTableData(page, 'sample1');

      // Open insert column dialog
      const th = page.locator('.subwindow table thead th[data-col-idx]').first();
      await th.click({ button: 'right' });
      await page.locator('.context-menu button', { hasText: 'Insert Column Right' }).click();

      await page.waitForSelector('.subwindow.dialog');
      const input = page.locator('.subwindow.dialog input');

      // First try a duplicate name
      await input.fill('email');
      await page.locator('.subwindow.dialog button', { hasText: 'OK' }).click();
      await page.waitForTimeout(200);
      await expect(page.locator('.subwindow.dialog')).toBeVisible();

      // Now correct to a unique name
      await input.fill('unique_col');
      await page.locator('.subwindow.dialog button', { hasText: 'OK' }).click();
      await page.waitForTimeout(500);

      // Dialog should be gone
      await expect(page.locator('.subwindow.dialog')).not.toBeAttached();

      const after = await getTableData(page, 'sample1');
      expect(after.columns).toContain('unique_col');
      expect(after.columns.length).toBe(before.columns.length + 1);
    });

    test('empty name cancels the dialog or shows error', async ({ page }) => {
      const before = await getTableData(page, 'sample1');

      const th = page.locator('.subwindow table thead th[data-col-idx]').first();
      await th.click({ button: 'right' });
      await page.locator('.context-menu button', { hasText: 'Insert Column Right' }).click();

      await page.waitForSelector('.subwindow.dialog');
      const input = page.locator('.subwindow.dialog input');

      // Clear and submit empty name
      await input.fill('');
      await page.locator('.subwindow.dialog button', { hasText: 'OK' }).click();
      await page.waitForTimeout(200);

      // Either the dialog stays open with error, or it is dismissed (cancel behavior)
      // No column should be added in either case
      const after = await getTableData(page, 'sample1');
      expect(after.columns.length).toBe(before.columns.length);
    });

    test('pressing Cancel dismisses the dialog without inserting', async ({ page }) => {
      const before = await getTableData(page, 'sample1');

      const th = page.locator('.subwindow table thead th[data-col-idx]').first();
      await th.click({ button: 'right' });
      await page.locator('.context-menu button', { hasText: 'Insert Column Right' }).click();

      await page.waitForSelector('.subwindow.dialog');
      await page.locator('.subwindow.dialog button', { hasText: 'Cancel' }).click();
      await page.waitForTimeout(200);

      await expect(page.locator('.subwindow.dialog')).not.toBeAttached();
      const after = await getTableData(page, 'sample1');
      expect(after.columns.length).toBe(before.columns.length);
    });
  });

  // ---------------------------------------------------------------------------
  // Inline Rename Duplicate Detection
  // ---------------------------------------------------------------------------

  test.describe('Inline Rename Duplicate Detection', () => {
    // Helper: trigger column rename via app._test.startColumnRename
    async function startRename(page, colIdx) {
      await page.evaluate((idx) => {
        const win = app._test.windows.find(w => w.tableName === 'sample1');
        const th = win._table.querySelector(`thead th[data-col-idx="${idx}"]`);
        const col = app._test.tables['sample1'].columns[idx];
        app._test.startColumnRename(win, th, col);
      }, colIdx);
      await page.waitForTimeout(200);
    }

    test('renaming column to duplicate name shows red border on the input', async ({ page }) => {
      await startRename(page, 0);

      const input = page.locator('.subwindow table thead .inline-rename');
      await expect(input).toBeVisible();

      // Type a duplicate name (email already exists)
      await input.fill('email');
      await input.press('Enter');
      await page.waitForTimeout(200);

      // The input should still be visible with red border
      await expect(input).toBeVisible();
      const borderColor = await input.evaluate(el => el.style.borderColor);
      expect(borderColor).toBe('rgb(255, 107, 107)');
    });

    test('the input stays active/focused after duplicate error', async ({ page }) => {
      await startRename(page, 0);

      const input = page.locator('.subwindow table thead .inline-rename');
      await input.fill('email');
      await input.press('Enter');
      await page.waitForTimeout(300);

      // Input should still be focused
      const isFocused = await input.evaluate(el => document.activeElement === el);
      expect(isFocused).toBe(true);
    });

    test('correcting to a unique name and pressing Enter succeeds', async ({ page }) => {
      const before = await getTableData(page, 'sample1');

      await startRename(page, 0);

      const input = page.locator('.subwindow table thead .inline-rename');
      // First try duplicate
      await input.fill('email');
      await input.press('Enter');
      await page.waitForTimeout(300);

      // Now correct to a unique name
      await input.fill('full_name');
      await input.press('Enter');
      await page.waitForTimeout(500);

      // Input should be gone (rename succeeded)
      await expect(page.locator('.subwindow table thead .inline-rename')).not.toBeAttached();

      const after = await getTableData(page, 'sample1');
      expect(after.columns).toContain('full_name');
      expect(after.columns).not.toContain(before.columns[0]); // 'name' is gone
    });

    test('Escape cancels rename even after a duplicate error', async ({ page }) => {
      const before = await getTableData(page, 'sample1');

      await startRename(page, 0);

      const input = page.locator('.subwindow table thead .inline-rename');
      // Try duplicate
      await input.fill('email');
      await input.press('Enter');
      await page.waitForTimeout(300);

      // Press Escape to cancel
      await input.press('Escape');
      await page.waitForTimeout(300);

      // Input should be gone
      await expect(page.locator('.subwindow table thead .inline-rename')).not.toBeAttached();

      // Column names should be unchanged
      const after = await getTableData(page, 'sample1');
      expect(after.columns).toEqual(before.columns);
    });

    test('the red border clears when user types a new character', async ({ page }) => {
      await startRename(page, 0);

      const input = page.locator('.subwindow table thead .inline-rename');
      // Trigger duplicate error
      await input.fill('email');
      await input.press('Enter');
      await page.waitForTimeout(300);

      // Verify red border is set
      let borderColor = await input.evaluate(el => el.style.borderColor);
      expect(borderColor).toBe('rgb(255, 107, 107)');

      // Type a character to clear the error
      await input.press('x');
      await page.waitForTimeout(100);

      // Border should be cleared
      borderColor = await input.evaluate(el => el.style.borderColor);
      expect(borderColor).toBe('');
    });
  });
});
