const path = require('path');
const os = require('os');
const { defineConfig } = require('@playwright/test');

module.exports = defineConfig({
  testDir: '.',
  testMatch: '**/*.spec.js',
  timeout: 60000,
  // Tests are fully independent (each opens its own page against the static
  // server), so they parallelize freely. The suite is sleep-bound (debounce
  // waits), so cores-2 beats Playwright's 50%-of-cores default.
  // Override with PW_WORKERS=1 to debug serially.
  workers: process.env.PW_WORKERS ? Number(process.env.PW_WORKERS) : Math.max(2, os.cpus().length - 2),
  fullyParallel: true,
  use: {
    baseURL: 'http://localhost:8274',
    headless: true,
    launchOptions: {
      args: ['--headless=new'],
    },
  },
  projects: [
    {
      name: 'chromium',
      use: {
        browserName: 'chromium',
        // Use the full chromium, not the headless shell
        channel: undefined,
      },
    },
  ],
  webServer: {
    command: 'python3 -m http.server 8274',
    port: 8274,
    cwd: path.resolve(__dirname, '..'),
    reuseExistingServer: true,
  },
});
