import { imageTagCount, stripImageTags } from './imageText.js';

// A4 한 장의 크기와 페이지 나누기 규칙.
// 화면 미리보기에서는 px, 인쇄(@page A4, 여백 0)에서는 같은 값이 그대로 A4에 맞습니다.

export const SHEET_W_MM = 210;
export const SHEET_H_MM = 297;

const MM_TO_PX = 96 / 25.4;

export const SHEET_W_PX = SHEET_W_MM * MM_TO_PX;
export const SHEET_H_PX = SHEET_H_MM * MM_TO_PX;

export const ROWS_PER_PAGE = 30;

const ROW_HEIGHT_MM = 8.2;
const LEARNING_MAIN_LINE_HEIGHT_MM = (9.75 * 1.15 * 25.4) / 72;
const LEARNING_KEY_LINE_HEIGHT_MM = (9.75 * 1.3 * 25.4) / 72;
const LEARNING_HINT_LINE_HEIGHT_MM = (7.25 * 1.1 * 25.4) / 72;
const LEARNING_KEY_TOP_PADDING_MM = 0.4;
const LEARNING_HINT_BOTTOM_PADDING_MM = 0.4;
const ROW_TOP_PADDING_MM = 0.75;
const ROW_BORDER_MM = 0.26;
const WIDE_CHARACTER_PATTERN = /\p{Script=Han}|\p{Script=Hangul}|\p{Script=Hiragana}|\p{Script=Katakana}/u;

function estimatedLineCount(value, charactersPerLine) {
  return String(value ?? '')
    .split(/\r\n?|\n/)
    .reduce(
      (total, line) => {
        const lineWidth = [...line].reduce(
          (width, character) => width + (WIDE_CHARACTER_PATTERN.test(character) ? 2 : 1),
          0
        );
        return total + Math.max(1, Math.ceil(lineWidth / charactersPerLine));
      },
      0
    );
}

// A4 화면에서 이미지는 CSS 로 이 높이까지만 보여 줍니다.
export const A4_IMAGE_MAX_HEIGHT_MM = 14;

function textHeightMm(value, charactersPerLine, lineHeightMm) {
  const images = imageTagCount(value);
  const plain = stripImageTags(value);
  const lines = images && !plain.trim() ? 0 : estimatedLineCount(plain, charactersPerLine);
  return lines * lineHeightMm + images * A4_IMAGE_MAX_HEIGHT_MM;
}

function sheetRowHeightMm(item) {
  const valueHeight = textHeightMm(item.value, 30, LEARNING_MAIN_LINE_HEIGHT_MM);
  const keyHeight = textHeightMm(item.key, 14, LEARNING_KEY_LINE_HEIGHT_MM) + LEARNING_KEY_TOP_PADDING_MM;
  const hintHeight = item.hint
    ? 0.4 + textHeightMm(item.hint, 54, LEARNING_HINT_LINE_HEIGHT_MM) + LEARNING_HINT_BOTTOM_PADDING_MM
    : 0;
  return Math.max(
    ROW_HEIGHT_MM,
    Math.max(keyHeight, valueHeight) + hintHeight +
      ROW_TOP_PADDING_MM +
      ROW_BORDER_MM
  );
}

export function shuffle(list) {
  const arr = [...list];
  for (let i = arr.length - 1; i > 0; i -= 1) {
    const j = Math.floor(Math.random() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
  return arr;
}

export function buildPages(items, size = ROWS_PER_PAGE) {
  const pages = [];
  const maxColumnHeight = size * ROW_HEIGHT_MM;
  let columns = [[], []];
  let columnHeights = [0, 0];

  items.forEach((item, index) => {
    const rowHeight = sheetRowHeightMm(item);

    let column = columnHeights[0] + rowHeight <= maxColumnHeight ? 0 : 1;

    if (columnHeights[column] + rowHeight > maxColumnHeight) {
      pages.push([...columns[0], ...columns[1]]);
      columns = [[], []];
      columnHeights = [0, 0];
      column = 0;
    }

    columns[column].push({
      ...item,
      index: index + 1,
      column,
      estimatedHeightMm: rowHeight,
    });
    columnHeights[column] += rowHeight;
  });

  if (columns[0].length || columns[1].length) {
    pages.push([...columns[0], ...columns[1]]);
  }

  const lastPage = pages.at(-1);
  if (
    lastPage &&
    lastPage.length > 1 &&
    (lastPage.every((item) => item.column === 0) || lastPage.every((item) => item.column === 1))
  ) {
    lastPage.sort((a, b) => a.index - b.index);
    const totalHeight = lastPage.reduce((sum, item) => sum + item.estimatedHeightMm, 0);
    let leftHeight = 0;
    let splitIndex = 1;
    let closestDifference = Infinity;
    for (let index = 1; index < lastPage.length; index += 1) {
      leftHeight += lastPage[index - 1].estimatedHeightMm;
      const difference = Math.abs(totalHeight - 2 * leftHeight);
      if (difference < closestDifference) {
        closestDifference = difference;
        splitIndex = index;
      }
    }
    lastPage.forEach((item, index) => {
      item.column = index < splitIndex ? 0 : 1;
    });
  }

  return pages;
}

export function sortByRecallProbability(items) {
  return items
    .map((item, index) => ({ item, index }))
    .sort((a, b) =>
      a.item.recallProbability - b.item.recallProbability || a.index - b.index
    )
    .map(({ item }) => item);
}

export function buildPriorityPages(reviewedItems, unreviewedItems, size = ROWS_PER_PAGE) {
  const reviewedPages = buildPages(reviewedItems, size).map((rows, index) => ({
    rows,
    section: 'reviewed',
    sectionStart: index === 0,
  }));
  const unreviewedPages = buildPages(unreviewedItems, size).map((rows, index) => ({
    rows: rows.map((row) => ({ ...row, index: row.index + reviewedItems.length })),
    section: 'unreviewed',
    sectionStart: index === 0,
  }));
  return [
    ...reviewedPages,
    ...unreviewedPages,
  ];
}