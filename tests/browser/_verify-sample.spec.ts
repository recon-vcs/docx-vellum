import { expect, test } from '@playwright/test';

interface DocxGlobal {
	renderSync(data: Blob, bodyContainer: HTMLElement, styleContainer: HTMLElement, options?: Record<string, unknown>): Promise<unknown>;
}

type HarnessWindow = Window & typeof globalThis & { docx: DocxGlobal };

test('verify zz-sample-analyze rendering after fixes', async ({ page }) => {
	const pageErrors: Error[] = [];
	page.on('pageerror', (error) => pageErrors.push(error));
	page.on('console', (msg) => {
		if (msg.type() === 'error') console.log('console.error:', msg.text());
	});

	await page.goto('/tests/browser/harness.html');
	const result = await page.evaluate(async () => {
		const response = await fetch('/tests/fixtures/zz-sample-analyze.docx');
		const blob = await response.blob();
		const body = document.querySelector<HTMLElement>('#document-container')!;
		const styleContainer = document.querySelector<HTMLElement>('#style-container')!;
		const docx = (window as unknown as HarnessWindow).docx;
		await docx.renderSync(blob, body, styleContainer, { breakPages: true });
		const sections = Array.from(document.querySelectorAll('section.docx-wrapper, section[class*="docx"]'));
		return {
			pageCount: sections.length,
			hasMathML: !!document.querySelector('math'),
			footerTexts: Array.from(document.querySelectorAll('footer')).map(f => f.textContent?.trim()),
			columnCounts: Array.from(document.querySelectorAll('article')).map(a => getComputedStyle(a).columnCount),
		};
	});

	console.log('RESULT:', JSON.stringify(result, null, 2));
	await page.screenshot({ path: 'test-results/zz-sample-full.png', fullPage: true });

	expect(pageErrors).toEqual([]);
});
