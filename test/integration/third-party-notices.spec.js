const { test, expect } = require('@playwright/test');
const { openApp } = require('../helpers');

// Every library bundled in lib/ must be credited in THIRD-PARTY-NOTICES.md.
const BUNDLED_LIBS = [
  'Papa Parse',
  'sql.js',
  'SheetJS',
  'JSZip',
  'Chart.js',
  'jsPDF',
  'jsPDF-AutoTable',
  'web-llm',
];

test.describe('Third-party license notices', () => {
  test('THIRD-PARTY-NOTICES.md is served and credits every bundled library', async ({ request }) => {
    const res = await request.get('/THIRD-PARTY-NOTICES.md');
    expect(res.status()).toBe(200);
    const text = await res.text();

    for (const lib of BUNDLED_LIBS) {
      expect(text, `missing credit for ${lib}`).toContain(lib);
    }
    // Full license texts must be present (Apache-2.0 for SheetJS/web-llm,
    // MIT for the rest), plus the JSZip dual-license election.
    expect(text).toContain('Apache License');
    expect(text).toContain('Version 2.0, January 2004');
    expect(text).toContain('MIT License');
    expect(text).toMatch(/JSZip is dual-licensed .* CSVSQL uses and\s*redistributes it under the MIT license/s);
  });

  test('PyPI static copies are in sync with root files', async ({ request }) => {
    // csvsql/static/ is a manual copy (see CLAUDE.md) — catch stale copies.
    for (const file of ['THIRD-PARTY-NOTICES.md', 'app.js', 'index.html', 'style.css']) {
      const root = await request.get(`/${file}`);
      const copy = await request.get(`/csvsql/static/${file}`);
      expect(root.status(), `/${file}`).toBe(200);
      expect(copy.status(), `/csvsql/static/${file}`).toBe(200);
      expect(await copy.text(), `${file} is stale in csvsql/static/`).toBe(await root.text());
    }
  });

  test('README credits the core libraries without naming AI-only ones', async ({ request }) => {
    const res = await request.get('/README.md');
    expect(res.status()).toBe(200);
    const text = await res.text();
    expect(text).toContain('THIRD-PARTY-NOTICES.md');
    for (const lib of ['Papa Parse', 'sql.js', 'SheetJS', 'JSZip']) {
      expect(text).toContain(lib);
    }
    // AI stays undocumented in README (owner's decision — see CLAUDE.md).
    expect(text.toLowerCase()).not.toContain('web-llm');
    expect(text.toLowerCase()).not.toContain('webllm');
  });

  test('About dialog credits libraries and points at the notices file', async ({ page }) => {
    await openApp(page);
    await page.evaluate(() => app.showAbout());
    const win = page.locator('.subwindow', { hasText: 'About CSVSQL' });
    await expect(win).toBeVisible();
    const text = await win.textContent();

    expect(text).toContain('Third-Party Libraries');
    expect(text).toContain('THIRD-PARTY-NOTICES.md');
    for (const lib of ['Papa Parse', 'sql.js', 'SheetJS', 'JSZip']) {
      expect(text).toContain(lib);
    }
    // AI-only dependencies must not be named in the UI.
    expect(text.toLowerCase()).not.toContain('web-llm');
    expect(text.toLowerCase()).not.toContain('webllm');
  });
});
