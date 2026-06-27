const { test, expect } = require('@playwright/test');
const { openApp } = require('../helpers');

// The prompt dialogs (New Table, Save As, Insert Column, …) and AI Settings are
// built on the reusable dialog subwindow component in modal mode: a dimmed
// backdrop, click-outside / Esc to dismiss, centered, and resizable.

test.describe('Modal dialogs', () => {
  // Ctrl+N opens the "New Table" name prompt — a modal dialog.
  async function openNewTablePrompt(page) {
    await page.keyboard.press('Control+n');
    await expect(page.locator('.subwindow.dialog')).toHaveCount(1);
  }

  test('modal prompt shows a backdrop over the workspace', async ({ page }) => {
    await openApp(page);
    await openNewTablePrompt(page);
    await expect(page.locator('.dialog-backdrop')).toHaveCount(1);
    await expect(page.locator('.dialog-backdrop')).toBeVisible();
  });

  test('clicking the backdrop cancels the prompt', async ({ page }) => {
    await openApp(page);
    await openNewTablePrompt(page);

    // Click near the top-left of the backdrop, away from the centered dialog.
    await page.locator('.dialog-backdrop').click({ position: { x: 5, y: 5 } });

    await expect(page.locator('.subwindow.dialog')).toHaveCount(0);
    await expect(page.locator('.dialog-backdrop')).toHaveCount(0);
    // No table was created.
    const count = await page.evaluate(() => app._test.windows.length);
    expect(count).toBe(0);
  });

  test('Escape cancels the prompt', async ({ page }) => {
    await openApp(page);
    await openNewTablePrompt(page);

    await page.keyboard.press('Escape');

    await expect(page.locator('.subwindow.dialog')).toHaveCount(0);
    await expect(page.locator('.dialog-backdrop')).toHaveCount(0);
    const count = await page.evaluate(() => app._test.windows.length);
    expect(count).toBe(0);
  });

  test('modal dialog is resizable (has resize handles)', async ({ page }) => {
    await openApp(page);
    await openNewTablePrompt(page);
    await expect(page.locator('.subwindow.dialog .resize-handle')).toHaveCount(8);
  });

  test('modal dialog has no minimize/maximize buttons or status bar', async ({ page }) => {
    await openApp(page);
    await openNewTablePrompt(page);
    await expect(page.locator('.subwindow.dialog .btn-min')).toHaveCount(0);
    await expect(page.locator('.subwindow.dialog .btn-max')).toHaveCount(0);
    await expect(page.locator('.subwindow.dialog .win-statusbar')).toHaveCount(0);
    await expect(page.locator('.subwindow.dialog .btn-close')).toHaveCount(1);
  });

  test('completing the prompt removes the backdrop and proceeds', async ({ page }) => {
    await openApp(page);
    await openNewTablePrompt(page);

    // First prompt: table name.
    await page.locator('.subwindow.dialog .modal-input').fill('modal_test');
    await page.locator('.subwindow.dialog .modal-buttons .ok').click();

    // Second prompt (column names) opens — still exactly one dialog + one backdrop.
    await expect(page.locator('.subwindow.dialog')).toHaveCount(1);
    await expect(page.locator('.dialog-backdrop')).toHaveCount(1);

    await page.locator('.subwindow.dialog .modal-input').fill('id, name');
    await page.locator('.subwindow.dialog .modal-buttons .ok').click();

    // Dialog and backdrop are gone; the table window exists.
    await expect(page.locator('.dialog-backdrop')).toHaveCount(0);
    const count = await page.evaluate(() => app._test.windows.length);
    expect(count).toBe(1);
  });

  test('modal dialog is horizontally centered in the workspace', async ({ page }) => {
    await openApp(page);
    await openNewTablePrompt(page);

    const centered = await page.evaluate(() => {
      const dlg = app._test.windows.find(w => w.isDialog);
      const area = document.getElementById('window-area');
      const left = parseInt(dlg.el.style.left);
      const w = dlg.el.offsetWidth;
      const expected = Math.round((area.clientWidth - w) / 2);
      return { left, expected };
    });
    // Allow a 1px rounding tolerance.
    expect(Math.abs(centered.left - centered.expected)).toBeLessThanOrEqual(1);
  });

  test('the legacy .modal-overlay / .modal DOM is no longer used', async ({ page }) => {
    await openApp(page);
    await openNewTablePrompt(page);
    await expect(page.locator('.modal-overlay')).toHaveCount(0);
    await expect(page.locator('.modal')).toHaveCount(0);
    // The prompt content still uses .modal-input / .modal-buttons inside the dialog.
    await expect(page.locator('.subwindow.dialog .modal-input')).toHaveCount(1);
  });

  test('a raw mousedown on the backdrop cancels the prompt', async ({ page }) => {
    await openApp(page);
    await openNewTablePrompt(page);

    // Dispatch a real mousedown (the backdrop closes on mousedown, not click).
    await page.locator('.dialog-backdrop').dispatchEvent('mousedown');

    await expect(page.locator('.subwindow.dialog')).toHaveCount(0);
    await expect(page.locator('.dialog-backdrop')).toHaveCount(0);
    const count = await page.evaluate(() => app._test.windows.length);
    expect(count).toBe(0);
  });

  test('showPrompt callback receives null on cancel and the value on confirm', async ({ page }) => {
    await openApp(page);

    // Cancel path → callback(null).
    const cancelled = await page.evaluate(() => new Promise((resolve) => {
      app._test.showPrompt('T', 'L', 'def', (v) => resolve(v));
      // Close via Escape on the next tick.
      setTimeout(() => document.dispatchEvent(
        new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })), 0);
    }));
    expect(cancelled).toBeNull();
    await expect(page.locator('.subwindow.dialog')).toHaveCount(0);

    // Confirm path → callback(value).
    const confirmed = await page.evaluate(() => new Promise((resolve) => {
      app._test.showPrompt('T', 'L', 'hello', (v) => resolve(v));
      setTimeout(() => {
        document.querySelector('.subwindow.dialog .modal-buttons .ok').click();
      }, 0);
    }));
    expect(confirmed).toBe('hello');
    await expect(page.locator('.subwindow.dialog')).toHaveCount(0);
  });
});

test.describe('AI Settings modal dialog', () => {
  async function openAISettings(page) {
    await page.evaluate(() => app._test.showAISettings());
    await expect(page.locator('.subwindow.dialog')).toHaveCount(1);
  }

  test('opens as a modal dialog with a backdrop', async ({ page }) => {
    await openApp(page);
    await openAISettings(page);
    await expect(page.locator('.dialog-backdrop')).toHaveCount(1);
    await expect(page.locator('.subwindow.dialog .win-title')).toHaveText('AI Settings');
    // Provider/model controls are present in the dialog body.
    await expect(page.locator('.subwindow.dialog #ai-set-provider')).toHaveCount(1);
  });

  test('is resizable with no min/max buttons or status bar', async ({ page }) => {
    await openApp(page);
    await openAISettings(page);
    await expect(page.locator('.subwindow.dialog .resize-handle')).toHaveCount(8);
    await expect(page.locator('.subwindow.dialog .btn-min')).toHaveCount(0);
    await expect(page.locator('.subwindow.dialog .btn-max')).toHaveCount(0);
    await expect(page.locator('.subwindow.dialog .win-statusbar')).toHaveCount(0);
  });

  test('Escape closes it without saving', async ({ page }) => {
    await openApp(page);
    await openAISettings(page);
    await page.keyboard.press('Escape');
    await expect(page.locator('.subwindow.dialog')).toHaveCount(0);
    await expect(page.locator('.dialog-backdrop')).toHaveCount(0);
  });

  test('Cancel button closes it', async ({ page }) => {
    await openApp(page);
    await openAISettings(page);
    await page.locator('.subwindow.dialog .modal-buttons .cancel').click();
    await expect(page.locator('.subwindow.dialog')).toHaveCount(0);
    await expect(page.locator('.dialog-backdrop')).toHaveCount(0);
  });

  test('clicking the backdrop closes it', async ({ page }) => {
    await openApp(page);
    await openAISettings(page);
    await page.locator('.dialog-backdrop').dispatchEvent('mousedown');
    await expect(page.locator('.subwindow.dialog')).toHaveCount(0);
    await expect(page.locator('.dialog-backdrop')).toHaveCount(0);
  });

  test('Save persists the selected provider', async ({ page }) => {
    await openApp(page);
    await openAISettings(page);
    await page.locator('.subwindow.dialog #ai-set-provider').selectOption('openai');
    await page.locator('.subwindow.dialog .modal-buttons .ok').click();
    await expect(page.locator('.subwindow.dialog')).toHaveCount(0);
    const provider = await page.evaluate(
      () => JSON.parse(localStorage.getItem('csvsql_ai_settings')).provider);
    expect(provider).toBe('openai');
  });
});
