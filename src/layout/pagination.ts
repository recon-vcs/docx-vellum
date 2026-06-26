import { OpenXmlElement } from '../model/element';
import { SectionProperties } from '../document/section';
import { LayoutRegion, PhysicalPage } from './layout-region';
import { buildPhysicalPages } from './page-builder';
import { buildSectionStream } from './section-stream';
import { splitRegionsByExplicitBreaks } from './explicit-breaks';

export interface PaginationPlan {
	regions: LayoutRegion[];
	pages: PhysicalPage[];
}

export function buildPaginationPlan(
	bodyChildren: OpenXmlElement[],
	rootSectProps: SectionProperties,
): PaginationPlan {
	const sectionRegions = buildSectionStream(bodyChildren, rootSectProps);
	const regions = splitRegionsByExplicitBreaks(sectionRegions);
	const pages = buildPhysicalPages(regions);

	return { regions, pages };
}
