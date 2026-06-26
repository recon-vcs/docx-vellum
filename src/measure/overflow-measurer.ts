export interface OverflowMetrics {
	clientHeight: number;
	scrollHeight: number;
}

export function isOverflowing(metrics: OverflowMetrics): boolean {
	return metrics.clientHeight < metrics.scrollHeight;
}

export function measureElementOverflow(el: HTMLElement): boolean {
	const currentOverflow = getComputedStyle(el).overflow;
	if (!currentOverflow || currentOverflow === 'visible') {
		el.style.overflow = 'hidden';
	}

	const overflow = isOverflowing({
		clientHeight: el.clientHeight,
		scrollHeight: el.scrollHeight,
	});

	el.style.overflow = currentOverflow;
	return overflow;
}
