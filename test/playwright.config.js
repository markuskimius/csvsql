const path = require('path');
const { defineConfig } = require('@playwright/test');

module.exports = defineConfig({
  testDir: '.',
  testMatch: '**/*.spec.js',
  timeout: 60000,
  workers: 1,
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
