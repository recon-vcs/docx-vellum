import { SectionType } from "./section";

export enum DomType {
	Document = "document",
	Page = "page",
	Paragraph = "paragraph",
	Run = "run",
	Break = "break",
	LastRenderedPageBreak = "lastRenderedPageBreak",
	SectionBreak = "sectionBreak",
	NoBreakHyphen = "noBreakHyphen",
	Table = "table",
	Row = "row",
	Cell = "cell",
	Hyperlink = "hyperlink",
	Drawing = "drawing",
	Image = "image",
	Text = "text",
	Character = "character",
	Tab = "tab",
	Symbol = "symbol",
	BookmarkStart = "bookmarkStart",
	BookmarkEnd = "bookmarkEnd",
	Footer = "footer",
	Header = "header",
	FootnoteReference = "footnoteReference",
	EndnoteReference = "endnoteReference",
	Footnotes = "footnotes",
	Footnote = "footnote",
	Endnotes = "endnotes",
	Endnote = "endnote",
	SimpleField = "simpleField",
	ComplexField = "complexField",
	Instruction = "instruction",
	VmlPicture = "vmlPicture",
	Shape = "shape",
	MmlMath = "mmlMath",
	MmlMathParagraph = "mmlMathParagraph",
	MmlFraction = "mmlFraction",
	MmlFunction = "mmlFunction",
	MmlFunctionName = "mmlFunctionName",
	MmlNumerator = "mmlNumerator",
	MmlDenominator = "mmlDenominator",
	MmlRadical = "mmlRadical",
	MmlBase = "mmlBase",
	MmlDegree = "mmlDegree",
	MmlSuperscript = "mmlSuperscript",
	MmlSubscript = "mmlSubscript",
	MmlPreSubSuper = "mmlPreSubSuper",
	MmlSubArgument = "mmlSubArgument",
	MmlSuperArgument = "mmlSuperArgument",
	MmlNary = "mmlNary",
	MmlDelimiter = "mmlDelimiter",
	MmlRun = "mmlRun",
	MmlEquationArray = "mmlEquationArray",
	MmlLimit = "mmlLimit",
	MmlLimitLower = "mmlLimitLower",
	MmlMatrix = "mmlMatrix",
	MmlMatrixRow = "mmlMatrixRow",
	MmlBox = "mmlBox",
	MmlBar = "mmlBar",
	MmlGroupChar = "mmlGroupChar",
	VmlElement = "vmlElement",
	Inserted = "inserted",
	Deleted = "deleted",
	DeletedText = "deletedText",
	Comment = "comment",
	CommentReference = "commentReference",
	CommentRangeStart = "commentRangeStart",
	CommentRangeEnd = "commentRangeEnd"
}

// TODO 分离Math类型
export enum MathDomType {
	Base = "mmlBase",
	Bar = "mmlBar",
	Box = "mmlBox",
	Delimiter = "mmlDelimiter",
	Degree = "mmlDegree",
	Denominator = "mmlDenominator",
	Function = "mmlFunction",
	FunctionName = "mmlFunctionName",
	Fraction = "mmlFraction",
	GroupChar = "mmlGroupChar",
	Limit = "mmlLimit",
	LimitLower = "mmlLimitLower",
	Matrix = "mmlMatrix",
	MatrixRow = "mmlMatrixRow",
	Math = "mmlMath",
	MathParagraph = "mmlMathParagraph",
	Nary = "mmlNary",
	Numerator = "mmlNumerator",
	PreSubSuper = "mmlPreSubSuper",
	Radical = "mmlRadical",
	SubArgument = "mmlSubArgument",
	Subscript = "mmlSubscript",
	Superscript = "mmlSuperscript",
}

export interface OpenXmlElement {
	uuid?: string,
	type: DomType;
	children?: OpenXmlElement[];
	cssStyle?: Record<string, any>;
	props?: Record<string, any>;
	// 元素层级
	level?: number;
	// 元素数组索引
	index?: number;
	// 溢出索引，数组，有多个元素溢出
	breakIndex?: Set<number>;
	//style name
	styleName?: string;
	//class mods
	className?: string;
	// 父级元素
	parent?: OpenXmlElement;
	sourcePath?: string;
}

export abstract class OpenXmlElementBase implements OpenXmlElement {
	type: DomType;
	children?: OpenXmlElement[] = [];
	cssStyle?: Record<string, any> = {};
	props?: Record<string, any>;
	// 元素层级
	level?: number;
	// 元素数组索引
	index?: number;
	// 溢出索引，数组，有多个元素溢出
	breakIndex?: Set<number>;
	//style name
	styleName?: string;
	//class mods
	className?: string;
	sourcePath?: string;

	parent?: OpenXmlElement;
}

export interface WmlHyperlink extends OpenXmlElement {
	id?: string;
	href?: string;
}

export interface WmlNoteReference extends OpenXmlElement {
	id: string;
}

export interface WmlBreak extends OpenXmlElement {
	break: BreakType;
}

export enum BreakType {
	Column = "column",
	Page = "page",
	// default
	TextWrapping = "textWrapping",
}

export interface WmlSectionBreak extends OpenXmlElement {
	break: SectionType;
}

export interface WmlLastRenderedPageBreak extends OpenXmlElement {

}

export interface WmlText extends OpenXmlElement {
	text: string;
}

export interface WmlCharacter extends OpenXmlElement {
	char: string;
}

export interface WmlSymbol extends OpenXmlElement {
	font: string;
	char: string;
}

export interface WmlTable extends OpenXmlElement {
	columns?: WmlTableColumn[];
	cellStyle?: Record<string, string>;

	colBandSize?: number;
	rowBandSize?: number;
}

export interface WmlTableRow extends OpenXmlElement {
	isHeader?: boolean;
}

export interface WmlTableCell extends OpenXmlElement {
	verticalMerge?: 'restart' | 'continue' | string;
	span?: number;
}

export interface WmlImage extends OpenXmlElement {
	src: string;
}

export enum WrapType {
	Inline = "Inline",
	None = "None",
	TopAndBottom = "TopAndBottom",
	Tight = "Tight",
	Through = "Through",
	Square = "Square",
	Polygon = "Polygon",
}

export interface WmlDrawing extends OpenXmlElement {

}

export interface WmlTableColumn {
	width?: string;
}

export interface IDomNumbering {
	id: string;
	level: number;
	start: number;
	pStyleName: string;
	pStyle: Record<string, string>;
	rStyle: Record<string, string>;
	levelText?: string;
	suff: string;
	format?: string;
	bullet?: NumberingPicBullet;
}

export interface NumberingPicBullet {
	id: number;
	src: string;
	style?: string;
}
