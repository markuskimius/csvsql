const { test, expect } = require('@playwright/test');
const { openApp } = require('../helpers');

test.describe('AI settings', () => {
  test('dialog lists all six providers', async ({ page }) => {
    await openApp(page);
    await page.evaluate(() => app._test.showAISettings());
    const providers = await page.evaluate(() =>
      [...document.querySelectorAll('#ai-set-provider option')].map(o => o.value));
    expect(providers).toEqual(['auto', 'claude', 'openai', 'gemini', 'grok', 'ollama', 'webllm']);
  });

  test('each cloud provider shows its model catalog and key field', async ({ page }) => {
    await openApp(page);
    await page.evaluate(() => app._test.showAISettings());

    const models = async () => page.evaluate(() =>
      [...document.querySelectorAll('#ai-set-model option')].map(o => o.value));
    const keyFieldVisible = async (id) => page.evaluate((sel) =>
      document.querySelector(sel).style.display !== 'none', `#ai-set-${id}-fields`);

    await page.selectOption('#ai-set-provider', 'claude');
    expect(await models()).toContain('claude-opus-4-8');
    expect(await models()).toContain('claude-sonnet-5');
    expect(await keyFieldVisible('claude')).toBe(true);
    expect(await keyFieldVisible('openai')).toBe(false);

    await page.selectOption('#ai-set-provider', 'openai');
    expect(await models()).toContain('gpt-5.1');
    expect(await keyFieldVisible('openai')).toBe(true);

    await page.selectOption('#ai-set-provider', 'gemini');
    expect(await models()).toContain('gemini-2.5-flash');
    expect(await keyFieldVisible('gemini')).toBe(true);

    await page.selectOption('#ai-set-provider', 'grok');
    expect(await models()).toContain('grok-4');
    expect(await keyFieldVisible('grok')).toBe(true);
  });

  test('retired saved model IDs migrate on load', async ({ page }) => {
    await page.addInitScript(() => {
      localStorage.setItem('csvsql_ai_settings', JSON.stringify({
        provider: 'claude',
        model: 'claude-opus-4-20250514', // retired June 2026
        ollamaUrl: 'http://localhost:11434',
        claudeApiKey: 'sk-ant-test',
        openaiApiKey: '',
      }));
    });
    await openApp(page);
    await page.evaluate(() => app._test.showAISettings());
    await expect(page.locator('#ai-set-model')).toHaveValue('claude-opus-4-8');
  });

  test('saving persists provider, model, and keys to localStorage', async ({ page }) => {
    await openApp(page);
    await page.evaluate(() => app._test.showAISettings());
    await page.selectOption('#ai-set-provider', 'gemini');
    await page.locator('#ai-set-gemini-key').fill('AIza-test-key');
    await page.selectOption('#ai-set-model', 'gemini-2.5-pro');
    await page.locator('.dialog-form .ok, .modal-buttons .ok').first().click();
    const saved = await page.evaluate(() => JSON.parse(localStorage.getItem('csvsql_ai_settings')));
    expect(saved.provider).toBe('gemini');
    expect(saved.model).toBe('gemini-2.5-pro');
    expect(saved.geminiApiKey).toBe('AIza-test-key');
  });
});
