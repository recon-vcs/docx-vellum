import { defineConfig } from 'vitest/config';

export default defineConfig({
	test: {
		// jsdom provides DOMParser/XML DOM used by the parsing pipeline.
		environment: 'jsdom',
		include: ['tests/unit/**/*.spec.ts'],
	},
});
