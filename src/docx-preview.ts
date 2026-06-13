import { WordDocument } from './word-document';
import { DocumentParser } from './document-parser';
// HTML renderer, asynchronous (legacy, inherited from docx-preview)
import { HtmlRenderer } from './html-renderer';
// HTML renderer, synchronous (pagination-aware)
import { HtmlRendererSync } from './html-renderer-sync';
import { createRenderResult, RenderResult } from './render-result';
export type { AttachOptions, OverlayHandle, OverlayLayer, PageHandle, RenderResult, SourceMap } from './render-result';

/** Accepted input formats for a docx document. */
export type DocumentSource = Blob | ArrayBuffer | Uint8Array;

export interface Options {
	breakPages: boolean;                    //enables page breaking on page breaks
	className: string;                      //class name/prefix for default and document style classes

	ignoreFonts: boolean;                   //disables fonts rendering
	ignoreHeight: boolean;                  //disables rendering height of page
	ignoreImageWrap: boolean;               //disables image text wrap setting
	ignoreLastRenderedPageBreak: boolean;   //disables page breaking on lastRenderedPageBreak elements
	ignoreTableWrap: boolean;               //disables table's text wrap setting
	ignoreWidth: boolean;                   //disables rendering width of page

	inWrapper: boolean;                     //enables rendering of wrapper around document content

	renderChanges: boolean;                 //enables experimental rendering of document changes (insertions/deletions)
	renderEndnotes: boolean;                //enables endnotes rendering
	renderFooters: boolean;                 //enables footers rendering
	renderFootnotes: boolean;               //enables footnotes rendering
	renderHeaders: boolean;                 //enables headers rendering

	trimXmlDeclaration: boolean;            //if true, xml declaration will be removed from xml documents before parsing
	useBase64URL: boolean;                  //if true, images, fonts, etc. will be converted to base 64 URL, otherwise URL.createObjectURL is used

	debug: boolean;                         //enables additional logging
	experimental: boolean;                  //enables experimental features (tab stops calculation)
}

export const defaultOptions: Options = {
	breakPages: true,
	className: "docx",

	ignoreFonts: false,
	ignoreHeight: false,
	ignoreImageWrap: false,
	ignoreLastRenderedPageBreak: true,
	ignoreTableWrap: true,
	ignoreWidth: false,

	inWrapper: true,

	renderChanges: false,
	renderEndnotes: true,
	renderFooters: true,
	renderFootnotes: true,
	renderHeaders: true,

	trimXmlDeclaration: true,
	useBase64URL: false,

	debug: false,
	experimental: false,
};

/** Parses a docx file into a WordDocument model without rendering it. */
export function parseAsync(data: DocumentSource, userOptions: Partial<Options> | null = null): Promise<WordDocument> {
	const ops: Options = { ...defaultOptions, ...userOptions };
	return WordDocument.load(data, new DocumentParser(ops), ops);
}

/** Renders an already-parsed WordDocument into the given containers. */
export async function renderDocument(
	document: WordDocument,
	bodyContainer: HTMLElement,
	styleContainer?: HTMLElement | null,
	sync: boolean = true,
	userOptions?: Partial<Options> | null,
): Promise<RenderResult> {
	const ops: Options = { ...defaultOptions, ...userOptions };
	const renderer = sync ? new HtmlRendererSync() : new HtmlRenderer();
	await renderer.render(document, bodyContainer, styleContainer ?? undefined, ops);
	return createRenderResult(document, bodyContainer, ops.className);
}

/** Parses and renders with the synchronous, pagination-aware renderer. */
export async function renderSync(
	data: DocumentSource,
	bodyContainer: HTMLElement,
	styleContainer: HTMLElement | null = null,
	userOptions: Partial<Options> | null = null,
): Promise<RenderResult> {
	const doc = await parseAsync(data, userOptions);
	return renderDocument(doc, bodyContainer, styleContainer, true, userOptions);
}

/** Parses and renders with the legacy asynchronous renderer. */
export async function renderAsync(
	data: DocumentSource,
	bodyContainer: HTMLElement,
	styleContainer?: HTMLElement | null,
	userOptions?: Partial<Options> | null,
): Promise<RenderResult> {
	const doc = await parseAsync(data, userOptions);
	return renderDocument(doc, bodyContainer, styleContainer, false, userOptions);
}
