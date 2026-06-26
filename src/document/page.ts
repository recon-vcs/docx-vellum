import { DomType, OpenXmlElement } from "./dom";
import { SectionProperties } from "./section";
import { uuid } from "../utils";
import type { LayoutRegion } from "../layout/layout-region";

export interface TreeNode extends OpenXmlElement {
	prev?: TreeNode | null;
	next?: TreeNode | null;
}

export interface PageProps {
	sectProps?: SectionProperties,
	children: OpenXmlElement[],
	stack?: TreeNode[],
	isSplit?: boolean,
	isFirstPage?: boolean;
	isLastPage?: boolean;
	breakIndex?: Set<number>;
	contentElement?: HTMLElement;
	checkingOverflow?: boolean,
	regions?: LayoutRegion[];
}

export class Page implements OpenXmlElement {
	type: DomType;
	pageId: string;
	sectProps?: SectionProperties;
	children: OpenXmlElement[];
	stack: TreeNode[];
	level?: number;
	isSplit: boolean;
	isFirstPage?: boolean;
	isLastPage?: boolean;
	breakIndex?: Set<number>;
	contentElement?: HTMLElement;
	checkingOverflow?: boolean;
	regions?: LayoutRegion[];

	constructor({ sectProps, children = [], stack = [], isSplit = false, isFirstPage = false, isLastPage = false, breakIndex = new Set(), contentElement, checkingOverflow = false, regions }: PageProps) {
		this.type = DomType.Page;
		this.level = 1;
		this.pageId = uuid();
		this.sectProps = sectProps;
		this.children = children;
		this.stack = stack;
		this.isSplit = isSplit;
		this.isFirstPage = isFirstPage;
		this.isLastPage = isLastPage;
		this.breakIndex = breakIndex;
		this.contentElement = contentElement;
		this.checkingOverflow = checkingOverflow;
		this.regions = regions;
	}
}
