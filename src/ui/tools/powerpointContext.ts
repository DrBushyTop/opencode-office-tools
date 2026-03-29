export interface PowerPointContextSnapshot {
  selectedSlideIds: string[];
  selectedShapeIds: string[];
  activeSlideId?: string;
  activeSlideIndex?: number;
}

let currentSnapshot: PowerPointContextSnapshot | null = null;

export function setPowerPointContextSnapshot(snapshot: PowerPointContextSnapshot | null) {
  currentSnapshot = snapshot ? { ...snapshot } : null;
}

export function getPowerPointContextSnapshot() {
  return currentSnapshot ? { ...currentSnapshot } : null;
}

export interface PowerPointTargetingArgs {
  slideIndex?: number;
  shapeId?: string | number;
  shapeIds?: Array<string | number>;
}

export function resolvePowerPointTargetingArgs<T extends PowerPointTargetingArgs>(args: T): T {
  const snapshot = getPowerPointContextSnapshot();
  if (!snapshot) return args;

  const next = { ...args };
  if (next.slideIndex === undefined && snapshot.activeSlideIndex !== undefined) {
    next.slideIndex = snapshot.activeSlideIndex;
  }

  if (next.shapeId === undefined && (!next.shapeIds || next.shapeIds.length === 0)) {
    if (snapshot.selectedShapeIds.length === 1) {
      next.shapeId = snapshot.selectedShapeIds[0];
    } else if (snapshot.selectedShapeIds.length > 1) {
      next.shapeIds = [...snapshot.selectedShapeIds];
    }
  }

  return next;
}

export function resolvePowerPointSlideIndexes(slideIndex: number | number[] | undefined): number | number[] | undefined {
  if (slideIndex !== undefined) return slideIndex;
  const snapshot = getPowerPointContextSnapshot();
  if (!snapshot) return undefined;
  if (snapshot.activeSlideIndex !== undefined) return snapshot.activeSlideIndex;
  if (snapshot.selectedSlideIds.length === 1 && snapshot.activeSlideIndex !== undefined) return snapshot.activeSlideIndex;
  return undefined;
}
