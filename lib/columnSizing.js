export function resizeAdjacentColumns(widths, leftIndex, delta, minWidth = 100) {
  if (
    !Array.isArray(widths) ||
    !Number.isInteger(leftIndex) ||
    leftIndex < 0 ||
    leftIndex >= widths.length - 1 ||
    !widths.every((width) => Number.isFinite(width) && width > 0) ||
    !Number.isFinite(delta) ||
    !Number.isFinite(minWidth) ||
    minWidth <= 0
  ) {
    throw new Error('Invalid column resize parameters.');
  }
  const left = widths[leftIndex];
  const right = widths[leftIndex + 1];
  const total = left + right;
  if (total < minWidth * 2) throw new Error('Adjacent columns are narrower than the minimum width.');
  const resizedLeft = Math.max(minWidth, Math.min(total - minWidth, left + delta));
  const resized = [...widths];
  resized[leftIndex] = resizedLeft;
  resized[leftIndex + 1] = total - resizedLeft;
  return resized;
}

export function normalizeColumnWidths(widths) {
  if (!Array.isArray(widths) || !widths.length || !widths.every((width) => Number.isFinite(width) && width > 0)) {
    throw new Error('Invalid column widths.');
  }
  const total = widths.reduce((sum, width) => sum + width, 0);
  return widths.map((width) => width / total);
}
