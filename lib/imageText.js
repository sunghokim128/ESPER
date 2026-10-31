// 셀 텍스트 안의 [img:ID] 참조를 다루는 순수 함수 모음입니다. 엑셀에는 이 참조만 저장됩니다.
const TAG_SOURCE = '\\[img:([A-Za-z0-9_-]{1,64})\\]';
const newTagPattern = () => new RegExp(TAG_SOURCE, 'g');

export const TRANSPARENT_PIXEL =
  'data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7';

export const imageTag = (id) => `[img:${id}]`;

export function extractImageIds(text) {
  return [...String(text ?? '').matchAll(newTagPattern())].map((match) => match[1]);
}

// 검색·음성·접근성 라벨처럼 글자만 필요한 곳에서는 태그를 텍스트로 취급하지 않습니다.
export function stripImageTags(text) {
  return String(text ?? '').replace(newTagPattern(), '');
}

export function imageTagCount(text) {
  return extractImageIds(text).length;
}

// 텍스트와 이미지를 원래 순서대로 나눈 조각 목록
export function splitImageText(text) {
  const source = String(text ?? '');
  const parts = [];
  let last = 0;
  for (const match of source.matchAll(newTagPattern())) {
    if (match.index > last) parts.push({ type: 'text', value: source.slice(last, match.index) });
    parts.push({ type: 'image', id: match[1] });
    last = match.index + match[0].length;
  }
  if (last < source.length) parts.push({ type: 'text', value: source.slice(last) });
  return parts;
}

export function imageElementHTML(id) {
  return `<img class="cell-img" data-img-id="${id}" src="${TRANSPARENT_PIXEL}" alt="">`;
}

export function renderImageText(text, escape) {
  return splitImageText(text)
    .map((part) => (part.type === 'image' ? imageElementHTML(part.id) : escape(part.value)))
    .join('');
}
