import { defineConfig } from '@playwright/test';
export default defineConfig({
  testDir: './tests/browser',
  fullyParallel: true,
  workers: 4,
  forbidOnly: !!process.env['CI'],
  retries: 0,
  reporter: [['list'], ['json', { outputFile: '.qa/browser-results.json' }]],
  use: { baseURL: 'http://127.0.0.1:41749', trace: 'retain-on-failure' },
  webServer: { command: 'node scripts/qa-browser-server.mjs', url: 'http://127.0.0.1:41749', reuseExistingServer: false },
  projects: ['chromium', 'firefox', 'webkit'].map(browserName => ({ name: browserName, use: { browserName: browserName as 'chromium' | 'firefox' | 'webkit' } })),
});
