import { readFile } from 'node:fs/promises';
import { expect, test, type Page } from '@playwright/test';

/**
 * Shape of the UMD global exposed by dist/docx-vellum.umd.js inside the
 * harness page. Only the members used by these tests are declared.
 */
interface DocxGlobal {
	renderAsync(data: Blob, bodyContainer: HTMLElement): Promise<unknown>;
	renderSync(data: Blob, bodyContainer: HTMLElement): Promise<unknown>;
}

type HarnessWindow = Window & typeof globalThis & { docx: DocxGlobal };

/**
 * Golden cases ported 1:1 from the upstream karma suite. The golden HTML in
 * tests/golden/ was captured from the asynchronous renderer with default
 * options and no style container, so the test replicates exactly that call.
 */
const GOLDEN_CASES = [
	'text',
	'underlines',
	'text-break',
	'table',
	'page-layout',
	'revision',
	'numbering',
	'line-spacing',
	'header-footer',
	'footnote',
	'equation',
] as const;

/** Representative fixtures for the synchronous renderer smoke tests. */
const SYNC_SMOKE_CASES = [
	'text',
	'table',
	'numbering',
	'header-footer',
	'footnote',
	'page-layout',
	'image-clip',
	'text-wrap-square',
	'comment',
	'break-page',
] as const;

/** Same whitespace normalization the upstream karma suite used. */
function normalizeHtml(text: string): string {
	return text.replace(/\t+|\s+/gi, ' ').replace(/></gi, '>\n<');
}

async function renderInPage(page: Page, fixture: string, mode: 'sync' | 'async'): Promise<string> {
	await page.goto('/tests/browser/harness.html');
	return page.evaluate(
		async ({ fixture, mode }) => {
			const response = await fetch(`/tests/fixtures/${fixture}.docx`);
			if (!response.ok) {
				throw new Error(`failed to fetch fixture ${fixture}: ${response.status}`);
			}
			const blob = await response.blob();
			const body = document.querySelector<HTMLElement>('#document-container');
			if (!body) {
				throw new Error('harness is missing #document-container');
			}
			const docx = (window as HarnessWindow).docx;
			if (mode === 'sync') {
				await docx.renderSync(blob, body);
			} else {
				await docx.renderAsync(blob, body);
			}
			return body.innerHTML;
		},
		{ fixture, mode },
	);
}

test.describe('renderAsync golden output', () => {
	for (const name of GOLDEN_CASES) {
		test(`renders ${name} matching golden HTML`, async ({ page }) => {
			const golden = await readFile(new URL(`../golden/${name}.html`, import.meta.url), 'utf-8');
			const actual = await renderInPage(page, name, 'async');
			expect(normalizeHtml(actual)).toBe(normalizeHtml(golden));
		});
	}
});

test.describe('renderSync smoke', () => {
	for (const name of SYNC_SMOKE_CASES) {
		test(`renders ${name} without errors`, async ({ page }) => {
			const pageErrors: Error[] = [];
			page.on('pageerror', (error) => pageErrors.push(error));

			const html = await renderInPage(page, name, 'sync');

			expect(pageErrors).toEqual([]);
			expect(html.length).toBeGreaterThan(0);
			expect(html).toContain('docx-wrapper');
		});
	}
});
