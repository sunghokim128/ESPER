import { createZipBlob, openZip, readSheetGrid } from './parseSheet.js';
import { extractImageIds } from './imageText.js';
import { isValidImageId } from './imageStore.js';

export const ESPR_EXTENSION = '.espr';
const ESPR_MIME = 'application/x-espr';
const IMAGE_DIRECTORY = 'images/';

export class EsprError extends Error {}

// 공유 파일에는 3열 엑셀과 그 안에서 실제로 쓰인 이미지만 담습니다. 학습 기록은 넣지 않습니다.
export async function createEsprBlob(workbookFile, getImageBlob) {
  const files = [{ name: workbookFile.name, data: new Uint8Array(await workbookFile.arrayBuffer()) }];
  const grid = await readSheetGrid(workbookFile);
  const ids = new Set(grid.flatMap((row) => row.flatMap((cell) => extractImageIds(cell))));
  let missingImages = 0;
  for (const id of ids) {
    const blob = await getImageBlob(id);
    if (!blob) {
      missingImages += 1;
      continue;
    }
    files.push({ name: `${IMAGE_DIRECTORY}${id}.jpg`, data: new Uint8Array(await blob.arrayBuffer()) });
  }
  return { blob: await createZipBlob(files, ESPR_MIME), imageCount: files.length - 1, missingImages };
}

export async function readEsprFile(file) {
  let zip;
  try {
    zip = await openZip(await file.arrayBuffer());
  } catch {
    throw new EsprError('invalid');
  }
  const workbookEntries = zip.entries.filter((entry) => /^[^/]+\.xlsx$/i.test(entry.name));
  if (workbookEntries.length !== 1) throw new EsprError('invalid');
  const workbookName = workbookEntries[0].name;
  const workbook = new File([await zip.readBytes(workbookName)], workbookName, {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const images = [];
  for (const entry of zip.entries) {
    const match = entry.name.match(/^images\/([^/]+)\.jpg$/i);
    if (!match || !isValidImageId(match[1])) continue;
    images.push({ id: match[1], blob: new Blob([await zip.readBytes(entry.name)], { type: 'image/jpeg' }) });
  }
  return { workbook, images };
}
