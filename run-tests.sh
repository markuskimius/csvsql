#!/bin/bash
cd "$(dirname "$0")"
npm install --silent 2>/dev/null
npx playwright install chromium 2>/dev/null
npx playwright test --config test/playwright.config.js "$@"
