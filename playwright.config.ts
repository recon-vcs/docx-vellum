import { defineConfig, devices } from '@playwright/test';

const PORT = 8765;

export default defineConfig({
	testDir: 'tests/browser',
	timeout: 60_000,
	fullyParallel: true,
	forbidOnly: !!process.env.CI,
	retries: 0,
	reporter: 'list',
	use: {
		baseURL: `http://127.0.0.1:${PORT}`,
	},
	projects: [
		{
			name: 'chromium',
			use: { ...devices['Desktop Chrome'] },
		},
	],
	webServer: {
		command: 'node tests/browser/static-server.mjs',
		url: `http://127.0.0.1:${PORT}/tests/browser/harness.html`,
		reuseExistingServer: !process.env.CI,
	},
});
