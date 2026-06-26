import { DocumentElement } from '../document/document';
import { SectionProperties } from '../document/section';
import { uuid } from '../utils';
import { splitRegionsByExplicitBreaks } from './explicit-breaks';
import { resolveHeaderFooterReferences } from './header-footer-context';
import { LayoutRegion, PhysicalPage } from './layout-region';
import { buildPhysicalPages } from './page-builder';
import { buildSectionStream } from './section-stream';

export interface ModernPageSplit {
	regions: LayoutRegion[];
	pages: PhysicalPage[];
}

export function splitDocumentIntoPhysicalPages(documentElement: DocumentElement): ModernPageSplit {
	const sectionRegions = normalizeRegionSections(
		buildSectionStream(documentElement.children ?? [], documentElement.sectProps),
	);
	const regions = splitRegionsByExplicitBreaks(sectionRegions);

	return {
		regions,
		pages: buildPhysicalPages(regions),
	};
}

function normalizeRegionSections(regions: LayoutRegion[]): LayoutRegion[] {
	let previousSection: SectionProperties | undefined;

	return regions.map((region) => {
		const section = normalizeSection(region.section, previousSection);
		previousSection = section;

		return {
			...region,
			section,
		};
	});
}

function normalizeSection(
	section: SectionProperties,
	previousSection: SectionProperties | undefined,
): SectionProperties {
	return {
		...section,
		sectionId: section.sectionId ?? uuid(),
		headerRefs: resolveHeaderFooterReferences(section.headerRefs, previousSection?.headerRefs),
		footerRefs: resolveHeaderFooterReferences(section.footerRefs, previousSection?.footerRefs),
	};
}
