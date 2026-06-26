import { WordDocument } from './word-document';
import { BreakType, DomType, IDomNumbering, OpenXmlElement, WmlBreak, WmlCharacter, WmlDrawing, WmlHyperlink, WmlImage, WmlLastRenderedPageBreak, WmlNoteReference, WmlSectionBreak, WmlSymbol, WmlTable, WmlTableCell, WmlTableColumn, WmlTableRow, WmlText, } from './document/dom';
import { CommonProperties } from './document/common';
import { Options } from './docx-preview';
import { DocumentElement } from './document/document';
import { WmlParagraph } from './document/paragraph';
import * as _ from 'lodash-es';
import { computePointToPixelRatio, updateTabStop } from './javascript';
import { FontTablePart } from './font-table/font-table';
import { FooterHeaderReference, SectionProperties } from './document/section';
import { Page, PageProps } from './document/page';
import { RunProperties, WmlRun } from './document/run';
import { WmlBookmarkStart } from './document/bookmarks';
import { WmlFieldSimple } from './document/fields';
import { IDomStyle } from './document/style';
import { WmlBaseNote, WmlEndnote, WmlEndnotes, WmlFootnote, WmlFootnotes } from './notes/elements';
import { ThemePart } from './theme/theme-part';
import { Part } from './common/part';
import { VmlElement } from './vml/vml';
import { WmlCommentRangeStart, WmlCommentReference } from './comments/elements';
import type { Stage } from 'konva/lib/Stage';
import type { Layer } from 'konva/lib/Layer';
import { Overflow, ChildrenType, createElement, createElementNS, removeAllElements, appendChildren, removeElements, createStyleElement, appendComment, checkOverflow, findParent } from './render/dom-utils';
import { TableContext, CellPos, CellVerticalMergeType, renderTable as renderTableFn, renderTableColumns as renderTableColumnsFn, renderTableRow as renderTableRowFn, renderTableCell as renderTableCellFn } from './render/table-renderer';
import { renderNotes as renderNotesFn, renderFootnoteReference as renderFootnoteReferenceFn, renderEndnoteReference as renderEndnoteReferenceFn } from './render/notes-renderer';
import { MathRendererCallbacks, mathJustificationToTextAlign, renderMmlMathParagraph as renderMmlMathParagraphFn, renderMmlRadical as renderMmlRadicalFn, renderMmlDelimiter as renderMmlDelimiterFn, renderMmlNary as renderMmlNaryFn, renderMmlPreSubSuper as renderMmlPreSubSuperFn, renderMmlGroupChar as renderMmlGroupCharFn, renderMmlBar as renderMmlBarFn, renderMmlRun as renderMmlRunFn, renderMllList as renderMllListFn } from './render/math-renderer';
import { DrawingRenderContext, createKonva as createKonvaFn, renderDrawing as renderDrawingFn, renderImage as renderImageFn, renderShape as renderShapeFn, renderVmlElement as renderVmlElementFn, renderVmlPicture as renderVmlPictureFn } from './render/drawing-renderer';
import { renderHeaderFooter as renderHeaderFooterFn } from './render/header-footer-renderer';
import { InlineRendererCallbacks, renderCharacter as renderCharacterFn, renderHyperlink as renderHyperlinkFn, renderParagraph as renderParagraphFn, renderRun as renderRunFn, renderText as renderTextFn } from './render/inline-renderer';
import { FieldsRendererCallbacks, renderBookmarkStart as renderBookmarkStartFn, renderCommentRangeEnd as renderCommentRangeEndFn, renderCommentRangeStart as renderCommentRangeStartFn, renderCommentReference as renderCommentReferenceFn, renderDeleted as renderDeletedFn, renderDeletedText as renderDeletedTextFn, renderInserted as renderInsertedFn, resolveFieldRuns as resolveFieldRunsFn, resolveSimpleField as resolveSimpleFieldFn } from './render/fields-renderer';
import { renderDefaultStyle as renderDefaultStyleFn, renderWrapper as renderWrapperFn } from './render/styles/default-styles';
import { processStyleName as processStyleNameFn, processStyles as processStylesFn, renderFontTable as renderFontTableFn, renderStyles as renderStylesFn, renderTheme as renderThemeFn } from './render/styles/document-styles';
import { levelTextToContent as levelTextToContentFn, numberingClass as numberingClassFn, numberingCounter as numberingCounterFn, numFormatToCssValue as numFormatToCssValueFn, processNumberings as processNumberingsFn, renderNumbering as renderNumberingFn, styleToString as styleToStringFn } from './render/styles/numbering-styles';
import { PageRendererCallbacks, createPage as createPageFn, createPageContent as createPageContentFn, renderHeaderFooterRef as renderHeaderFooterRefFn } from './render/page-renderer';
import { splitDocumentIntoPhysicalPages } from './layout/modern-page-splitter';
import { splitElementsByBreakIndex } from './render/split-by-break';

const ns = {
	html: 'http://www.w3.org/1999/xhtml',
	svg: 'http://www.w3.org/2000/svg',
	mathML: 'http://www.w3.org/1998/Math/MathML',
};

interface Node_DOM extends Node, Text {
	dataset: DOMStringMap;
}

// HTML渲染器
export class HtmlRendererSync {
	className = 'docx';
	rootSelector: string;
	document: WordDocument;
	options: Options;
	styleMap: Record<string, IDomStyle> = {};
	bodyContainer: HTMLElement;
	wrapper: HTMLElement;
	// 当前操作的Part
	currentPart: Part = null;
	// 系统的PPI
	pointToPixelRatio: number;

	// 当前操作的Page
	currentPage: Page;
	// Table rendering context (cell position + vertical-merge stacks for nested tables)
	tableCtx: TableContext = {
		tableVerticalMerges: [],
		currentVerticalMerge: null,
		tableCellPositions: [],
		currentCellPosition: null,
	};

	footnoteMap: Record<string, WmlFootnote> = {};
	endnoteMap: Record<string, WmlEndnote> = {};
	currentFootnoteIds: string[];
	currentEndnoteIds: string[] = [];
	// 已使用的Header、Footer部分的数组。
	usedHeaderFooterParts: any[] = [];

	defaultTabSize: string;
	// 当前制表位
	currentTabs: any[] = [];

	// Konva框架--stage元素
	konva_stage: Stage;
	// Konva框架--layer元素
	konva_layer: Layer;

	/**
	 * Object对象 => HTML标签
	 *
	 * @param document word文档Object对象
	 * @param bodyContainer HTML生成容器
	 * @param styleContainer CSS样式生成容器
	 * @param options 渲染配置选项
	 */

	async render(document: WordDocument, bodyContainer: HTMLElement, styleContainer: HTMLElement = null, options: Options) {
		// word文档对象
		this.document = document;
		// 渲染选项
		this.options = options;
		// class类前缀
		this.className = options.className;
		// 根元素
		this.rootSelector = options.inWrapper ? `.${this.className}-wrapper` : ':root';
		// 文档CSS样式
		this.styleMap = null;
		// 主体容器
		this.bodyContainer = bodyContainer;
		// 样式容器，可传参指定，默认为主体容器
		styleContainer = styleContainer || bodyContainer;
		// 计算Point/Pixel换算比例
		this.pointToPixelRatio = computePointToPixelRatio();
		// CSS样式生成容器，清空所有CSS样式
		removeAllElements(styleContainer);
		// HTML生成容器，清空所有HTML元素
		removeAllElements(bodyContainer);

		// 添加注释
		appendComment(styleContainer, 'docxjs library predefined styles');
		// 添加默认CSS样式
		styleContainer.appendChild(this.renderDefaultStyle());

		// 主题CSS样式
		if (document.themePart) {
			appendComment(styleContainer, 'docxjs document theme values');
			this.renderTheme(document.themePart, styleContainer);
		}
		// 文档默认CSS样式，包含表格、列表、段落、字体，样式存在继承顺序
		if (document.stylesPart != null) {
			this.styleMap = this.processStyles(document.stylesPart.styles);

			appendComment(styleContainer, 'docxjs document styles');
			styleContainer.appendChild(this.renderStyles(document.stylesPart.styles));
		}
		// 多级列表样式
		if (document.numberingPart) {
			this.processNumberings(document.numberingPart.domNumberings);

			appendComment(styleContainer, "docxjs document numbering styles");
			styleContainer.appendChild(this.renderNumbering(document.numberingPart.domNumberings, styleContainer));
			//styleContainer.appendChild(this.renderNumbering2(document.numberingPart, styleContainer));
		}
		// 字体列表CSS样式
		if (!options.ignoreFonts && document.fontTablePart) {
			this.renderFontTable(document.fontTablePart, styleContainer);
		}
		// 生成脚注部分的Map
		if (document.footnotesPart) {
			this.footnoteMap = _.keyBy(document.footnotesPart.rootElement.children, 'id');
		}
		// 生成尾注部分的Map
		if (document.endnotesPart) {
			this.endnoteMap = _.keyBy(document.endnotesPart.rootElement.children, 'id');
		}
		// 文档设置
		if (document.settingsPart) {
			this.defaultTabSize = document.settingsPart.settings?.defaultTabStop;
		}
		this.assignSourcePaths(document.documentPart.body.children);
		// 根据option生成wrapper
		if (this.options.inWrapper) {
			this.wrapper = this.renderWrapper();
			bodyContainer.appendChild(this.wrapper);
		} else {
			this.wrapper = bodyContainer;
		}
		// 生成Canvas画布元素--Konva框架
		this.renderKonva();
		// 主文档--内容
		await this.renderPages(document.documentPart.body);
		// 渲染完成所有Page, 隐藏Stage
		this.konva_stage.visible(false);
		// 刷新制表符
		this.refreshTabStops();
	}

	assignSourcePaths(children: OpenXmlElement[]) {
		children.forEach((child, index) => {
			const path = `body/${index}`;
			child.sourcePath = path;
			this.assignNestedSourcePaths(child, path);
		});
	}

	assignNestedSourcePaths(element: OpenXmlElement, path: string) {
		if (element.type === DomType.Table) {
			element.children?.forEach((row, rowIndex) => {
				row.children?.forEach((cell, cellIndex) => {
					const cellPath = `${path}/cell/${rowIndex}/${cellIndex}`;
					cell.sourcePath = cellPath;
					cell.children?.forEach((child) => {
						child.sourcePath = cellPath;
						this.assignNestedSourcePaths(child, cellPath);
					});
				});
			});
			return;
		}

		element.children?.forEach((child) => {
			child.sourcePath = path;
			this.assignNestedSourcePaths(child, path);
		});
	}

	// Render built-in default styles.
	renderDefaultStyle() {
		return renderDefaultStyleFn(this.className);
	}

	// Render document theme CSS variables.
	renderTheme(themePart: ThemePart, styleContainer: HTMLElement) {
		renderThemeFn(themePart, styleContainer, this.documentStylesCallbacks());
	}

	// Build a CSS class name for a Word style.
	processStyleName(className: string): string {
		return processStyleNameFn(className, this.className);
	}

	// Merge inherited style rules. Base styles are ordered before dependent styles.
	processStyles(styles: IDomStyle[]) {
		return processStylesFn(styles, this.documentStylesCallbacks());
	}

	// Render document style rules.
	renderStyles(styles: IDomStyle[]): HTMLElement {
		return renderStylesFn(styles, this.documentStylesCallbacks());
	}

	processNumberings(numberings: IDomNumbering[]) {
		processNumberingsFn(numberings, this.numberingStylesCallbacks());
	}

	renderNumbering(numberings: IDomNumbering[], styleContainer: HTMLElement) {
		return renderNumberingFn(numberings, styleContainer, this.numberingStylesCallbacks());
	}

	numberingClass(id: string, lvl: number) {
		return numberingClassFn(this.className, id, lvl);
	}

	styleToString(selectors: string, declarations: Record<string, string>, cssText: string = null) {
		return styleToStringFn(selectors, declarations, cssText);
	}

	numberingCounter(id: string, lvl: number) {
		return numberingCounterFn(this.className, id, lvl);
	}

	levelTextToContent(text: string, suff: string, id: string, numformat: string) {
		return levelTextToContentFn(text, suff, id, numformat, (counterId, level) => this.numberingCounter(counterId, level));
	}

	numFormatToCssValue(format: string) {
		return numFormatToCssValueFn(format);
	}

	// Render embedded font-face rules.
	renderFontTable(fontsPart: FontTablePart, styleContainer: HTMLElement) {
		renderFontTableFn(fontsPart, styleContainer, this.documentStylesCallbacks());
	}

	// Render the optional wrapper container.
	renderWrapper() {
		return renderWrapperFn(this.className);
	}

	private documentStylesCallbacks() {
		return {
			className: this.className,
			options: this.options,
			styleToString: (selectors, declarations, cssText = null) => this.styleToString(selectors, declarations, cssText),
			processStyleName: (className) => this.processStyleName(className),
			createStyleElement: (cssText) => createStyleElement(cssText),
			appendComment: (styleContainer, text) => appendComment(styleContainer, text),
			loadFont: (id, key) => this.document.loadFont(id, key),
			refreshTabStops: () => this.refreshTabStops(),
		};
	}

	private numberingStylesCallbacks() {
		return {
			className: this.className,
			rootSelector: this.rootSelector,
			findStyle: (styleName) => this.findStyle(styleName),
			styleToString: (selectors, declarations, cssText = null) => this.styleToString(selectors, declarations, cssText),
			createStyleElement: (cssText) => createStyleElement(cssText),
			loadNumberingImage: (src) => this.document.loadNumberingImage(src),
			numberingClass: (id, level) => this.numberingClass(id, level),
			numberingCounter: (id, level) => this.numberingCounter(id, level),
			levelTextToContent: (text, suff, id, numformat) => this.levelTextToContent(text, suff, id, numformat),
			numFormatToCssValue: (format) => this.numFormatToCssValue(format),
		};
	}

	// 复制CSS样式
	copyStyleProperties(input: Record<string, string>, output: Record<string, string>, attrs: string[] = null): Record<string, string> {
		if (!input) {
			return output;
		}
		if (output == null) {
			output = {};
		}
		if (attrs == null) {
			attrs = Object.getOwnPropertyNames(input);
		}

		for (const key of attrs) {
			if (input.hasOwnProperty(key) && !output.hasOwnProperty(key))
				output[key] = input[key];
		}

		return output;
	}

	// 递归明确元素parent父级关系
	processElement(element: OpenXmlElement) {
		if (element.children) {
			for (const e of element.children) {
				// 指向父级元素
				e.parent = element;
				// 标识其level层级
				e.level = element?.level + 1;
				// 判断类型
				if (e.type == DomType.Table) {
					// 处理表格style样式
					this.processTable(e);
					this.processElement(e);
				} else {
					// 递归渲染
					this.processElement(e);
				}
			}
		}
	}

	// 处理表格style样式
	processTable(table: WmlTable) {
		for (const r of table.children) {
			for (const c of r.children) {
				c.cssStyle = this.copyStyleProperties(table.cellStyle, c.cssStyle, [
					'border-left',
					'border-right',
					'border-top',
					'border-bottom',
					'padding-left',
					'padding-right',
					'padding-top',
					'padding-bottom',
				]);
			}
		}
	}

	// Build physical pages from section regions and explicit break data.
	splitPageBySymbol(documentElement: DocumentElement): Page[] {
		const split = splitDocumentIntoPhysicalPages(documentElement);

		return split.pages.map(physicalPage => {
			const activeRegion = physicalPage.regions[physicalPage.regions.length - 1];
			const children = physicalPage.regions.flatMap(region => region.children);

			return new Page({
				isSplit: false,
				sectProps: activeRegion?.section ?? documentElement.sectProps,
				children,
				regions: physicalPage.regions,
			} as PageProps);
		});
	}

	// 生成所有的页面Page
	async renderPages(document: DocumentElement) {
		// 根据options.breakPages，选择是否分页
		let pages: Page[];
		if (this.options.breakPages) {
			// 根据分页符，初步拆分页面
			pages = this.splitPageBySymbol(document);
		} else {
			// 不分页则，只有一个page
			pages = [new Page({ isSplit: true, sectProps: document.sectProps, children: document.children, } as PageProps)];
		}
		// 初步分页结果,缓存至body中
		document.pages = pages;
		// 前一个节属性，判断分节符的第一个page
		let prevProps = null;
		// 浅拷贝初步分页结果，后续拆分操作将不断扩充数组，导致下面循环异常
		let origin_pages = [...pages];
		// 遍历生成每一个page
		for (let i = 0; i < origin_pages.length; i++) {
			this.currentFootnoteIds = [];
			const page: Page = origin_pages[i];
			const { sectProps } = page;
			// sectionProps属性不存在，则使用文档级别props;
			page.sectProps = sectProps ?? document.sectProps;
			// 是否本小节的第一个page
			page.isFirstPage = prevProps != page.sectProps;
			// TODO 是否最后一个page,此时分页未完成，计算并不准确，影响到尾注的渲染
			page.isLastPage = i === origin_pages.length - 1;
			// 溢出检测默认不开启
			page.checkingOverflow = false;
			// 将上述数据存储在currentPage中
			this.currentPage = page;
			// 存储前一个节属性
			prevProps = page.sectProps;
			// 渲染单个page
			await this.renderPage();
		}
	}

	// Render a single page. If content overflows, recursively splits into the next page.
	async renderPage() {
		const { pageId, sectProps, children, isFirstPage, isLastPage, regions } = this.currentPage;
		this.processElement(this.currentPage);
		const pageElement = this.createPage(this.className, sectProps);

		this.renderStyleValues(this.document.documentPart.body.cssStyle, pageElement);

		const pages = this.document.documentPart.body.pages;
		const pageIndex = pages.findIndex((page) => page.pageId === pageId);

		let oHeader: HTMLElement = null;
		let oFooter: HTMLElement = null;
		if (this.options.renderHeaders) {
			oHeader = await this.renderHeaderFooterRef(sectProps.headerRefs, sectProps, pageIndex, isFirstPage, pageElement);
		}
		if (this.options.renderFooters) {
			oFooter = await this.renderHeaderFooterRef(sectProps.footerRefs, sectProps, pageIndex, isFirstPage, pageElement);
		}

		const getOffsetHeight = (el: HTMLElement) => (el?.offsetHeight ?? 0) * this.pointToPixelRatio;
		const { pageSize, pageMargins } = sectProps;
		const headerHeight = getOffsetHeight(oHeader);
		const footerHeight = getOffsetHeight(oFooter);
		const actualTop = _.max([parseFloat(pageMargins.top), headerHeight]);
		const actualBottom = _.max([parseFloat(pageMargins.bottom), footerHeight]);
		pageElement.style.paddingTop = `${actualTop}pt`;
		pageElement.style.paddingBottom = `${actualBottom}pt`;
		const contentHeight = parseFloat(pageSize.height) - actualTop - actualBottom;

		if (regions && regions.length > 1) {
			// Multi-region page: continuous section changes within one physical page.
			// Each region gets its own article with region-specific column layout.
			// Overflow detection is skipped; pre-computed layout is used as-is.
			for (const region of regions) {
				const regionArticle = this.createPageContent(region.section);
				regionArticle.style.minHeight = `${contentHeight}pt`;
				pageElement.appendChild(regionArticle);
				this.currentPage.contentElement = regionArticle;
				this.currentPage.checkingOverflow = false;
				await this.renderElements(region.children, regionArticle);
			}
			this.currentPage.isSplit = true;
			pages[pageIndex] = this.currentPage;
		} else {
			// Single-region page: standard overflow detection path.
			const contentElement = this.createPageContent(sectProps);
			if (this.options.breakPages) {
				contentElement.style.height = `${contentHeight}pt`;
			} else {
				contentElement.style.minHeight = `${contentHeight}pt`;
			}
			this.currentPage.contentElement = contentElement;
			pageElement.appendChild(contentElement);
			this.currentPage.checkingOverflow = true;
			const is_overflow = await this.renderElements(children, contentElement);
			if (is_overflow === Overflow.FALSE) {
				this.currentPage.isSplit = true;
				pages[pageIndex] = this.currentPage;
			}
			this.currentPage.checkingOverflow = false;
		}

		if (this.options.renderFootnotes) {
			await this.renderNotes(DomType.Footnotes, this.currentFootnoteIds, this.footnoteMap, pageElement);
		}
		if (this.options.renderEndnotes && isLastPage) {
			await this.renderNotes(DomType.Endnotes, this.currentEndnoteIds, this.endnoteMap, pageElement);
		}
	}

	// 创建Page
	createPage(className: string, props: SectionProperties) {
		return createPageFn(className, props, this.wrapper, {
			ignoreWidth: this.options.ignoreWidth,
			ignoreHeight: this.options.ignoreHeight,
		});
	}

	// TODO 分栏：一个页面可能存在多个章节section，每个section拥有不同的分栏
	// 多列分栏布局
	createPageContent(props: SectionProperties): HTMLElement {
		return createPageContentFn(props);
	}

	// TODO 分页不准确，页脚页码混乱，
	// TODO 支持奇数页偶数页不同页眉页脚
	// 渲染页眉/页脚的Ref
	async renderHeaderFooterRef(refs: FooterHeaderReference[], props: SectionProperties, pageIndex: number, isFirstPage: boolean, parent: HTMLElement) {
		return renderHeaderFooterRefFn(refs, props, pageIndex, isFirstPage, parent, this.pageRendererCallbacks());
	}

	private pageRendererCallbacks(): PageRendererCallbacks {
		return {
			document: this.document,
			ignoreWidth: this.options.ignoreWidth,
			ignoreHeight: this.options.ignoreHeight,
			evenAndOddHeaders: this.document.settingsPart.settings.evenAndOddHeaders,
			usedHeaderFooterParts: this.usedHeaderFooterParts,
			setCurrentPart: part => { this.currentPart = part; },
			processElement: elem => this.processElement(elem),
			renderHeaderFooter: (elem, tagName, parent) => this.renderHeaderFooter(elem, tagName, parent),
		};
	}

	// TODO 字体太大，尾注位置不对
	// 渲染脚注/尾注
	async renderNotes(type: DomType = DomType.Footnotes, noteIds: string[], notesMap: Record<string, WmlBaseNote>, parent: HTMLElement) {
		return renderNotesFn(type, noteIds, notesMap, parent, {
			processElement: (e) => this.processElement(e),
			renderChildren: (e, p) => this.renderChildren(e, p),
		});
	}

	// 根据XML对象渲染出多元素
	async renderElements(children: OpenXmlElement[], parent: HTMLElement | Element | Text): Promise<Overflow> {
		// 子元素溢出状态的数组
		let overflows: Overflow[] = [];
		// 已拆分的Pages数组
		let pages: Page[] = this.document.documentPart.body.pages;
		// 当前Page
		let { pageId, sectProps, children: current_page_children } = this.currentPage;

		// 计算当前Page的索引
		let pageIndex: number = pages.findIndex((page) => page.pageId === pageId);

		for (let i = 0; i < children.length; i++) {
			const elem = children[i];
			// 标识元素的索引
			elem.index = i;
			// 子元素溢出索引数组
			if (!elem.breakIndex) {
				elem.breakIndex = new Set();
			}
			// 根据XML对象渲染单个元素
			const rendered_element = await this.renderElement(elem, parent);
			// elem元素是否溢出
			let overflow: Overflow = rendered_element?.dataset?.overflow as Overflow ?? Overflow.UNKNOWN;
			// 下一步操作，终止循环/跳过此次遍历，进入下一次遍历
			let action: string;

			switch (overflow) {
				// 元素自身溢出
				case Overflow.SELF:
					// 缓存溢出元素的索引至自身的breakIndex.
					elem.breakIndex.add(0);
					// 缓存溢出元素的索引至父级的breakIndex。
					elem.parent.breakIndex.add(i);
					// 删除溢出元素
					removeElements(rendered_element, parent);
					action = 'break';
					break;

				// 叶子元素溢出
				case Overflow.TRUE:
				// 插入元素children之后，全部child溢出
				case Overflow.FULL:
					// 缓存溢出元素的索引至父级的breakIndex。
					elem.parent.breakIndex.add(i);
					// 删除溢出元素
					if (elem.type !== DomType.Cell) {
						removeElements(rendered_element, parent);
					}
					action = 'break';
					break;

				// 插入元素children之后，一部分child溢出
				case Overflow.PART:
					// 缓存溢出元素的索引至父级的breakIndex。
					elem.parent.breakIndex.add(i);
					action = 'break';
					break;

				// 未溢出
				case Overflow.FALSE:
				// 未执行溢出检测
				case Overflow.UNKNOWN:
				// 忽略溢出检测
				case Overflow.IGNORE:
					action = 'continue';
					break;

				default:
					action = 'continue';
					if (this.options.debug) {
						console.error('unhandled overflow', overflow, elem);
					}
			}
			// TableRow中存在多个td溢出
			if (elem.type === DomType.Cell) {
				action = 'continue';
			}
			// 将elem元素的溢出状态保存至数组
			overflows.push(overflow);
			// 跳过此次遍历，进入下一次遍历，后续代码不执行；
			if (action === 'continue') {
				continue;
			}
			// 顶层元素：溢出，action === break
			if (elem.level === 2) {
				// 根据breakIndex索引，删除后续元素，原始数组保留前面已经渲染的元素
				let next_page_children: OpenXmlElement[] = current_page_children.splice(i);
				// 生成新的page，新Page的sectionProps沿用前一页的sectionProps
				const next_page: Page = new Page({ sectProps, children: next_page_children } as PageProps);
				splitElementsByBreakIndex(this.currentPage, next_page);
				// 修改当前Page的状态
				this.currentPage.isSplit = true;
				this.currentPage.checkingOverflow = false;
				// 重新递归建立元素的parent父级关系
				this.processElement(this.currentPage);
				// 缓存当前page至pages
				pages[pageIndex] = this.currentPage;
				// 缓存拆分出去的新page
				pages.splice(pageIndex + 1, 0, next_page);
				// 新Page覆盖current_page的属性
				this.currentPage = next_page;
				// 重启新一个page的渲染
				await this.renderPage();
			}
			// 终止循环
			break;
		}
		/*
		* 推断elem父级元素溢出类型，overflows数组由于上述循环break的影响，后续子元素溢出状态不会存在，可能只有一个值。
		* 推断规则如下：
		* [Overflow.FULL,..., Overflow.TRUE,Overflow.SELF,Overflow.PART]：全部子元素溢出，推断溢出类型为Overflow.FULL;
		* [Overflow.PART,Overflow.PART,Overflow.PART]：所有子元素部分溢出，推断溢出类型为Overflow.PART;
		* [Overflow.FALSE,..., Overflow.TRUE,Overflow.IGNORE]：部分子元素溢出，推断溢出类型为Overflow.PART;
		* [Overflow.FALSE,Overflow.UNKNOWN,Overflow.IGNORE]：所有元素未溢出Overflow.FALSE，推断溢出类型为Overflow.FALSE;
		* [Overflow.UNKNOWN,Overflow.UNKNOWN,Overflow.UNKNOWN]：所有元素未知Overflow.UNKNOWN，推断溢出类型为Overflow.UNKNOWN;
		*
		* 注意，表格中，Row元素推断溢出类型必须遍历所有子元素。
		*/
		// 如果没有子元素或数组为空，则返回FALSE。注意，every遍历空数组返回true。
		if (overflows.length === 0) {
			return Overflow.FALSE;
		}
		// 溢出状态集合
		let overflowStatus: Overflow[] = [Overflow.FULL, Overflow.SELF, Overflow.TRUE, Overflow.PART];
		// 所有子元素部分溢出，推断溢出类型为Overflow.PART;
		let isAllPart: boolean = overflows.every(overflow => overflow === Overflow.PART);
		if (isAllPart) {
			return Overflow.PART;
		}
		// 是否全溢出
		let isFull: boolean = overflows.every(overflow => overflowStatus.includes(overflow));
		if (isFull) {
			return Overflow.FULL;
		}
		// 是否未执行溢出检测
		let isUnknown: boolean = overflows.every(overflow => overflow === Overflow.UNKNOWN);
		if (isUnknown) {
			return Overflow.UNKNOWN;
		}
		// 是否未溢出
		let isFalse: boolean = overflows.every(overflow => [Overflow.FALSE, Overflow.UNKNOWN, Overflow.IGNORE].includes(overflow));
		if (isFalse) {
			return Overflow.FALSE;
		}
		// 是否部分溢出
		let isPart: boolean = overflows.some(overflow => overflowStatus.includes(overflow));
		if (isPart) {
			return Overflow.PART;
		}
		return Overflow.UNKNOWN;
	}

	// 根据XML对象渲染单个元素
	async renderElement(elem: OpenXmlElement, parent?: HTMLElement | Element | Text): Promise<Node_DOM> {
		let oNode;

		switch (elem.type) {
			case DomType.Paragraph:
				oNode = await this.renderParagraph(elem as WmlParagraph, parent as HTMLElement);
				break;

			case DomType.Run:
				oNode = await this.renderRun(elem as WmlRun, parent as HTMLElement);
				break;

			case DomType.SimpleField:
				// container has no visual representation of its own; its
				// resolved Run(s) are appended directly to parent.
				await this.renderElements(
					this.resolveSimpleField(elem as WmlFieldSimple),
					parent as HTMLElement
				);
				oNode = null;
				break;

			case DomType.Text:
				oNode = await this.renderText(elem as WmlText, parent as HTMLElement);
				break;

			case DomType.Character:
				oNode = await this.renderCharacter(elem as WmlCharacter, parent as Text);
				break;

			case DomType.Table:
				oNode = await this.renderTable(elem as WmlTable, parent as HTMLElement);
				break;

			case DomType.Row:
				oNode = await this.renderTableRow(elem as WmlTableRow, parent as HTMLElement);
				break;

			case DomType.Cell:
				oNode = await this.renderTableCell(elem as WmlTableCell, parent as HTMLElement);
				break;

			case DomType.Hyperlink:
				oNode = await this.renderHyperlink(elem, parent as HTMLElement);
				break;

			case DomType.Drawing:
				oNode = await renderDrawingFn(elem as WmlDrawing, parent as HTMLElement, this.drawingRenderContext());
				break;

			case DomType.Image:
				oNode = await renderImageFn(elem as WmlImage, parent as HTMLElement, this.drawingRenderContext());
				break;

			case DomType.Shape:
				oNode = await renderShapeFn(elem, parent as HTMLElement, this.drawingRenderContext());
				break;

			case DomType.BookmarkStart:
				oNode = renderBookmarkStartFn(elem as WmlBookmarkStart, parent as HTMLElement, this.fieldsCallbacks());
				break;

			case DomType.BookmarkEnd:
				//ignore bookmark end
				oNode = null;
				break;

			case DomType.Tab:
				oNode = await this.renderTab(elem, parent as HTMLElement);
				break;

			case DomType.Symbol:
				oNode = await this.renderSymbol(elem as WmlSymbol, parent as HTMLElement);
				break;

			case DomType.Break:
				oNode = await this.renderBreak(elem as WmlBreak, parent as HTMLElement);
				break;

			case DomType.LastRenderedPageBreak:
				oNode = await this.renderLastRenderedPageBreak(elem as WmlLastRenderedPageBreak, parent as HTMLElement);
				break;

			case DomType.SectionBreak:
				oNode = await this.renderSectionBreak(elem as WmlSectionBreak, parent as HTMLElement);
				break;

			case DomType.Inserted:
				oNode = await renderInsertedFn(elem, parent as HTMLElement, this.fieldsCallbacks());
				break;

			case DomType.Deleted:
				oNode = await renderDeletedFn(elem, parent as HTMLElement, this.fieldsCallbacks());
				break;

			case DomType.DeletedText:
				oNode = await renderDeletedTextFn(elem as WmlText, parent as HTMLElement, this.fieldsCallbacks());
				break;

			case DomType.NoBreakHyphen:
				oNode = createElement('wbr');
				if (parent) {
					await this.appendChildren(parent as HTMLElement, oNode);
				}
				break;

			case DomType.CommentRangeStart:
				oNode = renderCommentRangeStartFn(elem as WmlCommentRangeStart);
				if (parent) appendChildren(parent, oNode);
				break;

			case DomType.CommentRangeEnd:
				oNode = renderCommentRangeEndFn(elem as WmlCommentRangeStart);
				if (parent) appendChildren(parent, oNode);
				break;

			case DomType.CommentReference:
				oNode = renderCommentReferenceFn(elem as WmlCommentReference, this.fieldsCallbacks());
				if (parent) appendChildren(parent, oNode);
				break;

			case DomType.Footer:
				oNode = await this.renderHeaderFooter(elem, 'footer', parent as HTMLElement);
				break;

			case DomType.Header:
				oNode = await this.renderHeaderFooter(elem, 'header', parent as HTMLElement);
				break;

			case DomType.Footnote:
			case DomType.Endnote:
				oNode = await this.renderContainer(elem, 'li');
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.FootnoteReference:
				oNode = renderFootnoteReferenceFn(elem as WmlNoteReference, this.currentFootnoteIds);
				if (parent) appendChildren(parent, oNode);
				break;

			case DomType.EndnoteReference:
				oNode = renderEndnoteReferenceFn(elem as WmlNoteReference, this.currentEndnoteIds);
				if (parent) appendChildren(parent, oNode);
				break;

			case DomType.VmlElement:
				oNode = await renderVmlElementFn(elem as VmlElement, parent as HTMLElement, this.drawingRenderContext());
				break;

			case DomType.VmlPicture:
				oNode = await renderVmlPictureFn(elem, this.drawingRenderContext());
				if (parent) appendChildren(parent, oNode);
				break;

			case DomType.MmlMath:
				oNode = await this.renderContainerNS(elem, ns.mathML, 'math', {
					xmlns: ns.mathML,
				});
				// TODO 作为子元素插入,针对此元素进行溢出检测
				if (parent) {
					oNode.dataset.overflow = await this.appendChildren(parent as HTMLElement, oNode);
				}
				break;

			case DomType.MmlMathParagraph:
				oNode = await renderMmlMathParagraphFn(elem, this.mathCallbacks());
				if (parent) appendChildren(parent, oNode);
				break;

			case DomType.MmlFraction:
				oNode = await this.renderContainerNS(elem, ns.mathML, 'mfrac');
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.MmlBase:
				oNode = await this.renderContainerNS(elem, ns.mathML, elem.parent.type == DomType.MmlMatrixRow ? "mtd" : "mrow");
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.MmlNumerator:
			case DomType.MmlDenominator:
			case DomType.MmlFunction:
			case DomType.MmlLimit:
			case DomType.MmlBox:
				oNode = await this.renderContainerNS(elem, ns.mathML, 'mrow');
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.MmlGroupChar:
				oNode = await renderMmlGroupCharFn(elem, this.mathCallbacks());
				if (parent) appendChildren(parent, oNode);
				break;

			case DomType.MmlLimitLower:
				oNode = await this.renderContainerNS(elem, ns.mathML, 'munder');
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.MmlMatrix:
				oNode = await this.renderContainerNS(elem, ns.mathML, 'mtable');
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.MmlMatrixRow:
				oNode = await this.renderContainerNS(elem, ns.mathML, 'mtr');
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.MmlRadical:
				oNode = await renderMmlRadicalFn(elem, this.mathCallbacks());
				if (parent) appendChildren(parent, oNode);
				break;

			case DomType.MmlSuperscript:
				oNode = await this.renderContainerNS(elem, ns.mathML, 'msup');
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.MmlSubscript:
				oNode = await this.renderContainerNS(elem, ns.mathML, 'msub');
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.MmlDegree:
			case DomType.MmlSuperArgument:
			case DomType.MmlSubArgument:
				oNode = await this.renderContainerNS(elem, ns.mathML, 'mn');
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.MmlFunctionName:
				oNode = await this.renderContainerNS(elem, ns.mathML, 'ms');
				// 作为子元素插入,忽略溢出检测
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.MmlDelimiter:
				oNode = await renderMmlDelimiterFn(elem, this.mathCallbacks());
				if (parent) appendChildren(parent, oNode);
				break;

			case DomType.MmlRun:
				oNode = await renderMmlRunFn(elem, this.mathCallbacks());
				if (parent) appendChildren(parent, oNode);
				break;

			case DomType.MmlNary:
				oNode = await renderMmlNaryFn(elem, this.mathCallbacks());
				if (parent) appendChildren(parent, oNode);
				break;

			case DomType.MmlPreSubSuper:
				oNode = await renderMmlPreSubSuperFn(elem, this.mathCallbacks());
				if (parent) {
					appendChildren(parent, oNode);
				}
				break;

			case DomType.MmlBar:
				oNode = await renderMmlBarFn(elem, this.mathCallbacks());
				if (parent) appendChildren(parent, oNode);
				break;

			case DomType.MmlEquationArray:
				oNode = await renderMllListFn(elem, this.mathCallbacks());
				if (parent) appendChildren(parent, oNode);
				break;
		}
		// 标记其XML标签名
		if (oNode && oNode?.nodeType === 1) {
			oNode.dataset.tag = elem.type;
			if (elem.sourcePath && this.isSourceAnchor(elem)) {
				oNode.dataset.vellumPath = elem.sourcePath;
			}
		}

		return oNode;
	}

	isSourceAnchor(elem: OpenXmlElement): boolean {
		return elem.type === DomType.Paragraph || elem.type === DomType.Table || elem.type === DomType.Cell || elem.type === DomType.SectionBreak;
	}

	// 根据XML对象渲染子元素，并插入父级元素
	async renderChildren(elem: OpenXmlElement, parent: HTMLElement | Element | Text): Promise<Overflow> {
		return await this.renderElements(elem.children, parent);
	}

	// 插入子元素，针对后代元素进行溢出检测
	async appendChildren(parent: HTMLElement | Text, children: ChildrenType): Promise<Overflow> {
		// 插入元素
		appendChildren(parent, children);

		let { isSplit, contentElement, checkingOverflow, } = this.currentPage;
		// 当前page已拆分，忽略溢出检测
		if (isSplit) {
			return Overflow.UNKNOWN;
		}
		// 当前page未拆分，是否需要溢出检测
		if (checkingOverflow) {
			// 溢出检测
			let isOverflow = checkOverflow(contentElement);
			return isOverflow ? Overflow.TRUE : Overflow.FALSE;
		} else {
			return Overflow.UNKNOWN;
		}
	}

	async renderContainer(elem: OpenXmlElement, tagName: keyof HTMLElementTagNameMap, props?: Record<string, any>) {
		const oContainer = createElement(tagName, props);

		oContainer.dataset.overflow = await this.renderChildren(elem, oContainer);
		return oContainer;
	}

	async renderContainerNS(elem: OpenXmlElement, ns: string, tagName: string, props?: Record<string, any>) {
		const parent = createElementNS(ns, tagName as any, props);
		await this.renderChildren(elem, parent);
		return parent;
	}

	mathJustificationToTextAlign(justification?: string) {
		return mathJustificationToTextAlign(justification);
	}

	private mathCallbacks(): MathRendererCallbacks {
		return {
			renderElement: (e, p) => this.renderElement(e, p as HTMLElement) as any,
			renderElements: (es, p) => this.renderElements(es, p as HTMLElement),
			renderChildren: (e, p) => this.renderChildren(e, p as HTMLElement),
			renderContainerNS: (e, n, t, pr) => this.renderContainerNS(e, n, t, pr),
			renderClass: (e, o) => this.renderClass(e, o as HTMLElement),
			renderStyleValues: (s, o) => this.renderStyleValues(s, o),
			className: this.className,
		};
	}

	resolveFieldRuns(runs: OpenXmlElement[]): OpenXmlElement[] {
		return resolveFieldRunsFn(runs, this.fieldsCallbacks());
	}

	resolveSimpleField(elem: WmlFieldSimple): OpenXmlElement[] {
		return resolveSimpleFieldFn(elem, this.fieldsCallbacks());
	}

	private fieldsCallbacks(): FieldsRendererCallbacks {
		return {
			processElement: (e) => this.processElement(e),
			renderChildren: (e, p) => this.renderChildren(e, p),
			appendChildren: (p, c) => this.appendChildren(p, c),
			appendChildrenWithoutOverflow: (p, c) => appendChildren(p, c),
			renderText: (e, p) => this.renderText(e, p),
			findComment: (id) => this.document.commentsPart?.commentMap[id],
			currentPageNumber: () => {
				const pages = this.document.documentPart.body.pages ?? [];
				return pages.findIndex(p => p.pageId === this.currentPage.pageId) + 1;
			},
			pageCount: () => this.document.documentPart.body.pages?.length ?? 0,
			renderChanges: () => this.options.renderChanges,
		};
	}

	private inlineCallbacks(): InlineRendererCallbacks {
		return {
			appendChildren: (p, c) => this.appendChildren(p, c),
			renderChildren: (e, p) => this.renderChildren(e, p),
			renderClass: (e, o) => this.renderClass(e, o),
			renderCommonProperties: (s, p) => this.renderCommonProperties(s, p),
			renderStyleValues: (s, o) => this.renderStyleValues(s, o),
			resolveFieldRuns: (runs) => this.resolveFieldRuns(runs),
			findStyle: (styleName) => this.findStyle(styleName),
			numberingClass: (id, level) => this.numberingClass(id, level),
			currentPageIsSplit: () => this.currentPage.isSplit,
			currentSectionProperties: () => this.currentPage.sectProps,
			findExternalRelation: (id) => this.document.documentPart.rels.find(
				it => it.id == id && it.targetMode === 'External'
			),
		};
	}

	async renderParagraph(elem: WmlParagraph, parent: HTMLElement) {
		return renderParagraphFn(elem, parent, this.inlineCallbacks());
	}

	async renderRun(elem: WmlRun, parent: HTMLElement) {
		return renderRunFn(elem, parent, this.inlineCallbacks());
	}

	async renderText(elem: WmlText, parent: HTMLElement) {
		return renderTextFn(elem, parent, this.inlineCallbacks());
	}

	// Render one character at a time for overflow detection.
	async renderCharacter(elem: WmlCharacter, parent: Text) {
		return renderCharacterFn(elem, parent, this.inlineCallbacks());
	}

	async renderTable(elem: WmlTable, parent: HTMLElement) {
		return renderTableFn(elem, parent, this.tableCtx, {
			renderChildren: (e, p) => this.renderChildren(e, p),
			appendChildren: (p, c) => this.appendChildren(p, c),
			renderClass: (e, o) => this.renderClass(e, o),
			renderStyleValues: (s, o) => this.renderStyleValues(s, o),
		});
	}

	// 表格--列
	renderTableColumns(columns: WmlTableColumn[], parent: HTMLElement) {
		return renderTableColumnsFn(columns, parent);
	}

	// 表格--行
	async renderTableRow(elem: OpenXmlElement, parent: HTMLElement) {
		return renderTableRowFn(elem, parent, this.tableCtx, {
			renderChildren: (e, p) => this.renderChildren(e, p),
			appendChildren: (p, c) => this.appendChildren(p, c),
			renderClass: (e, o) => this.renderClass(e, o),
			renderStyleValues: (s, o) => this.renderStyleValues(s, o),
		});
	}

	// 表格--单元格
	async renderTableCell(elem: WmlTableCell, parent: HTMLElement) {
		return renderTableCellFn(elem, parent, this.tableCtx, {
			renderChildren: (e, p) => this.renderChildren(e, p),
			appendChildren: (p, c) => this.appendChildren(p, c),
			renderClass: (e, o) => this.renderClass(e, o),
			renderStyleValues: (s, o) => this.renderStyleValues(s, o),
		});
	}

	async renderHyperlink(elem: WmlHyperlink, parent: HTMLElement) {
		return renderHyperlinkFn(elem, parent, this.inlineCallbacks());
	}

	renderKonva() {
		const { stage, layer } = createKonvaFn(this.bodyContainer);
		this.konva_stage = stage;
		this.konva_layer = layer;
	}

	private drawingRenderContext(): DrawingRenderContext {
		return {
			document: this.document,
			currentPart: this.currentPart,
			options: this.options,
			konvaStage: this.konva_stage,
			konvaLayer: this.konva_layer,
			appendChildren: (p, c) => this.appendChildren(p, c),
			renderChildren: (e, p) => this.renderChildren(e, p),
			renderElement: (e, p) => this.renderElement(e, p),
			renderStyleValues: (s, o) => this.renderStyleValues(s, o),
		};
	}

	// 渲染制表符
	async renderTab(elem: OpenXmlElement, parent: HTMLElement) {
		const tabSpan = createElement('span');

		tabSpan.innerHTML = '&nbsp;';

		tabSpan.className = this.tabStopClass();
		const stops = findParent<WmlParagraph>(elem, DomType.Paragraph).props?.tabs;
		this.currentTabs.push({ stops, span: tabSpan });

		// 作为子元素插入，执行溢出检测
		if (parent) {
			await this.appendChildren(parent, tabSpan);
		}

		return tabSpan;
	}

	async renderSymbol(elem: WmlSymbol, parent: HTMLElement) {
		const oSymbol = createElement('span');
		oSymbol.style.fontFamily = elem.font;
		oSymbol.innerHTML = `&#x${elem.char};`;
		// 溢出标识
		let is_overflow: Overflow;
		// oSymbol作为子元素插入，针对此元素进行溢出检测
		is_overflow = await this.appendChildren(parent, oSymbol);

		if (is_overflow === Overflow.TRUE) {
			oSymbol.dataset.overflow = Overflow.SELF;
		}

		oSymbol.dataset.overflow = is_overflow;

		return oSymbol;
	}

	// 渲染换行符号
	async renderBreak(elem: WmlBreak, parent: HTMLElement) {
		let oBreak: HTMLElement;

		switch (elem.break) {
			// 分页符
			case BreakType.Page:
				oBreak = createElement('br');
				// 添加class
				oBreak.classList.add('break', 'page');
				break;

			// 	TODO 分栏符
			case BreakType.Column:
				oBreak = createElement('br');
				// 添加class
				oBreak.classList.add('break', 'column');
				break;

			// 强制换行
			case BreakType.TextWrapping:
			default:
				oBreak = createElement('br');
				// 添加class
				oBreak.classList.add('break', 'textWrap');
				break;
		}
		// oBreak作为子元素插入，针对此元素执行溢出检测
		let isOverflow = await this.appendChildren(parent, oBreak);

		if (isOverflow === Overflow.TRUE) {
			isOverflow = Overflow.SELF;
		}

		oBreak.dataset.overflow = isOverflow;

		return oBreak;
	}

	async renderLastRenderedPageBreak(elem: WmlLastRenderedPageBreak, parent: HTMLElement) {
		const oLastRenderedPageBreak = createElement('wbr');
		// 添加class
		oLastRenderedPageBreak.classList.add('lastRenderedPageBreak');
		// oLastRenderedPageBreak作为子元素插入，针对此元素执行溢出检测
		let isOverflow = await this.appendChildren(parent, oLastRenderedPageBreak);
		// if true,empty element should be Overflow.SELF
		if (isOverflow === Overflow.TRUE) {
			isOverflow = Overflow.SELF;
		}

		oLastRenderedPageBreak.dataset.overflow = isOverflow;

		return oLastRenderedPageBreak;
	}

	async renderSectionBreak(elem: WmlSectionBreak, parent: HTMLElement) {
		const oSectionBreak = createElement('s');
		// 添加class
		oSectionBreak.classList.add('break', 'section');
		// oSectionBreak作为子元素插入，针对此元素执行溢出检测
		let isOverflow = await this.appendChildren(parent, oSectionBreak);
		// if true,empty element should be Overflow.SELF
		if (isOverflow === Overflow.TRUE) {
			isOverflow = Overflow.SELF;
		}

		oSectionBreak.dataset.overflow = isOverflow;
		// break type
		oSectionBreak.dataset.type = elem.break;

		return oSectionBreak;
	}

	// 渲染页眉页脚
	async renderHeaderFooter(elem: OpenXmlElement, tagName: keyof HTMLElementTagNameMap, parent: HTMLElement) {
		return renderHeaderFooterFn(elem, tagName, parent, {
			renderChildren: (e, p) => this.renderChildren(e, p),
			renderStyleValues: (s, o) => this.renderStyleValues(s, o),
		});
	}


	// 设置元素style样式
	renderStyleValues(style: Record<string, string>, output: HTMLElement) {
		for (const k in style) {
			if (k.startsWith('$')) {
				output.setAttribute(k.slice(1), style[k]);
			} else {
				output.style[k] = style[k];
			}
		}
	}

	renderRunProperties(style: any, props: RunProperties) {
		this.renderCommonProperties(style, props);
	}

	renderCommonProperties(style: any, props: CommonProperties) {
		if (props == null) return;

		if (props.color) {
			style['color'] = props.color;
		}

		if (props.fontSize) {
			style['font-size'] = props.fontSize;
		}
	}

	// 添加class类名
	renderClass(input: OpenXmlElement, output: HTMLElement | Element) {
		if (input.className) {
			output.className = input.className;
		}

		if (input.styleName) {
			output.classList.add(this.processStyleName(input.styleName));
		}
	}

	// 查找内置默认style样式
	findStyle(styleName: string) {
		return styleName && this.styleMap?.[styleName];
	}

	tabStopClass() {
		return `${this.className}-tab-stop`;
	}

	// 刷新tab制表符
	refreshTabStops() {
		for (const tab of this.currentTabs) {
			updateTabStop(tab.span, tab.stops, this.defaultTabSize, this.pointToPixelRatio);
		}
	}
}

// DOM utility functions moved to src/render/dom-utils.ts
