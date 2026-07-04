const { test, expect } = require('@playwright/test');
const { openApp } = require('../helpers');

// Open the app with Shift held through the load, then move the mouse while
// still holding it — the gesture that unlocks the hidden AI tab (the page can
// only see a held Shift via the modifier flag on an input event).
async function openAppWithShift(page) {
  await page.keyboard.down('Shift');
  await openApp(page);
  await page.waitForLoadState('load');
  await page.mouse.move(200, 200);
  await page.keyboard.up('Shift');
}

test.describe('AI unlock', () => {
  test('AI tab is hidden by default and the welcome message does not mention AI', async ({ page }) => {
    await openApp(page);
    expect(await page.evaluate(() => app._test.aiUnlocked)).toBe(false);
    await expect(page.locator('#console-tab-ai')).toBeHidden();
    const text = await page.locator('.empty-state-privacy').textContent();
    expect(text).toContain('never sent to any server');
    expect(text).not.toContain('AI');
  });

  test('Shift held through load reveals the AI tab and updates the welcome message', async ({ page }) => {
    await openAppWithShift(page);
    await expect(page.locator('#console-tab-ai')).toBeVisible();
    expect(await page.evaluate(() => app._test.aiUnlocked)).toBe(true);
    const text = await page.locator('.empty-state-privacy').textContent();
    expect(text).toContain('unless you use an AI model');
  });

  test('reloading with Shift held unlocks deterministically, regardless of release timing', async ({ page }) => {
    await openApp(page);
    // Hold Shift on the outgoing page (as during Ctrl+Shift+R), then reload:
    // the pagehide flag must unlock the new page with no further input
    await page.keyboard.down('Shift');
    await page.reload();
    await page.waitForFunction(() => window._appReady === true);
    await page.keyboard.up('Shift'); // release order/timing no longer matters
    await expect(page.locator('#console-tab-ai')).toBeVisible();
    expect(await page.evaluate(() => app._test.aiUnlocked)).toBe(true);
  });

  test('other keys pressed at startup do not end detection', async ({ page }) => {
    await page.keyboard.down('Shift');
    await openApp(page);
    await page.waitForLoadState('load');
    // Shift-less key events from other keys (e.g. releasing the rest of a
    // reload combo) must be ignored, not treated as proof Shift is up
    await page.evaluate(() => {
      window.dispatchEvent(new KeyboardEvent('keyup', { key: 'Meta', bubbles: true }));
      window.dispatchEvent(new KeyboardEvent('keydown', { key: 'F5', bubbles: true }));
    });
    await page.mouse.move(200, 200); // Shift still held → unlocks
    await page.keyboard.up('Shift');
    await expect(page.locator('#console-tab-ai')).toBeVisible();
    expect(await page.evaluate(() => app._test.aiUnlocked)).toBe(true);
  });

  test('releasing a held Shift without other input does not unlock', async ({ page }) => {
    await page.keyboard.down('Shift');
    await openApp(page);
    await page.waitForLoadState('load');
    await page.keyboard.up('Shift'); // release is the first event: no unlock, detection ends
    await page.waitForTimeout(100);
    await expect(page.locator('#console-tab-ai')).toBeHidden();
    expect(await page.evaluate(() => app._test.aiUnlocked)).toBe(false);
  });

  test('pressing Shift after the page has loaded does nothing', async ({ page }) => {
    await openApp(page);
    await page.waitForLoadState('load');
    await page.keyboard.press('Shift');
    await page.waitForTimeout(100);
    await expect(page.locator('#console-tab-ai')).toBeHidden();
    expect(await page.evaluate(() => app._test.aiUnlocked)).toBe(false);
  });

  test('once detection has ended, even Shift-held input does not unlock', async ({ page }) => {
    await openApp(page);
    await page.waitForLoadState('load');
    await page.mouse.move(100, 100); // first input event, no Shift → detection over
    await page.keyboard.down('Shift');
    await page.mouse.move(150, 150); // Shift genuinely held now, but too late
    await page.keyboard.up('Shift');
    await page.waitForTimeout(100);
    await expect(page.locator('#console-tab-ai')).toBeHidden();
    expect(await page.evaluate(() => app._test.aiUnlocked)).toBe(false);
  });

  test('unlocked AI tab opens the chat UI', async ({ page }) => {
    await openAppWithShift(page);
    await page.locator('#console-tab-ai').click();
    await expect(page.locator('#ai-body')).toBeVisible();
    await expect(page.locator('#ai-input-row')).toBeVisible();
    await expect(page.locator('#ai-controls')).toBeVisible();
  });

  test('the unlock is not persisted across reloads', async ({ page }) => {
    await openAppWithShift(page);
    await expect(page.locator('#console-tab-ai')).toBeVisible();

    await openApp(page);
    await expect(page.locator('#console-tab-ai')).toBeHidden();
    expect(await page.evaluate(() => app._test.aiUnlocked)).toBe(false);
  });

  test('while locked, no AI-related network requests are made', async ({ page }) => {
    await page.addInitScript(() => {
      // The provider that would probe localhost:11434 if AI were reachable
      localStorage.setItem('csvsql_ai_settings', JSON.stringify({
        provider: 'ollama', model: '',
        ollamaUrl: 'http://localhost:11434',
        claudeApiKey: '', openaiApiKey: '', geminiApiKey: '', grokApiKey: '',
      }));
    });
    await openApp(page);

    const requests = [];
    page.on('request', r => {
      const url = new URL(r.url());
      // Ignore the app's own static assets from the test server
      if (url.port !== '8274') requests.push(r.url());
    });
    await page.waitForTimeout(400);
    await expect(page.locator('#console-tab-ai')).toBeHidden();
    expect(requests).toEqual([]);
  });
});
