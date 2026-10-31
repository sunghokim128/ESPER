export const MAX_IMAGE_SIDE = 800;
const JPEG_QUALITY = 0.82;

const IMAGE_FOLDER_NAME = 'images';
const urlCache = new Map();
let directoryProvider = async () => null;

export function configureImageDirectory(provider) {
  directoryProvider = provider;
}

async function imageDirectory(projectName, create) {
  if (!projectName) throw new Error('이미지를 저장하거나 불러올 프로젝트가 지정되지 않았어요.');
  const root = await directoryProvider();
  if (!root) throw new Error('저장 폴더에 접근할 수 없어 이미지를 처리하지 못했어요.');
  let projectDirectory;
  try {
    projectDirectory = await root.getDirectoryHandle(projectName, { create });
  } catch (error) {
    if (!create && error instanceof DOMException && error.name === 'NotFoundError') return null;
    throw error;
  }
  try {
    return await projectDirectory.getDirectoryHandle(IMAGE_FOLDER_NAME, { create });
  } catch (error) {
    if (!create && error instanceof DOMException && error.name === 'NotFoundError') return null;
    throw error;
  }
}

async function writeImageFile(directory, id, blob) {
  const handle = await directory.getFileHandle(`${id}.jpg`, { create: true });
  const writable = await handle.createWritable();
  try {
    await writable.write(blob);
    await writable.close();
  } catch (error) {
    try {
      await writable.abort();
    } catch (abortError) {
      throw new AggregateError([error, abortError], '이미지 저장에 실패했고 임시 파일도 취소하지 못했습니다.');
    }
    throw error;
  }
}

async function readImageFile(directory, id) {
  try {
    return await (await directory.getFileHandle(`${id}.jpg`)).getFile();
  } catch (error) {
    if (error instanceof DOMException && error.name === 'NotFoundError') return null;
    throw error;
  }
}

async function loadElement(blob) {
  const url = URL.createObjectURL(blob);
  try {
    const image = new Image();
    image.src = url;
    await image.decode();
    return image;
  } finally {
    URL.revokeObjectURL(url);
  }
}

async function hashBytes(buffer) {
  const digest = await crypto.subtle.digest('SHA-256', buffer);
  return [...new Uint8Array(digest)]
    .slice(0, 12)
    .map((byte) => byte.toString(16).padStart(2, '0'))
    .join('');
}

export async function processImage(blob, maxSide = MAX_IMAGE_SIDE) {
  let source;
  let width;
  let height;
  try {
    source = await createImageBitmap(blob);
    width = source.width;
    height = source.height;
  } catch {
    source = await loadElement(blob);
    width = source.naturalWidth || 300;
    height = source.naturalHeight || 150;
  }
  if (!width || !height) throw new Error('empty image');
  const scale = Math.min(1, maxSide / Math.max(width, height));
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.round(width * scale));
  canvas.height = Math.max(1, Math.round(height * scale));
  const context = canvas.getContext('2d');
  context.fillStyle = '#fff';
  context.fillRect(0, 0, canvas.width, canvas.height);
  context.drawImage(source, 0, 0, canvas.width, canvas.height);
  source.close?.();
  const processed = await new Promise((resolve, reject) =>
    canvas.toBlob((result) => (result ? resolve(result) : reject(new Error('encode failed'))), 'image/jpeg', JPEG_QUALITY)
  );
  const id = await hashBytes(await processed.arrayBuffer());
  return { id, blob: processed };
}

export const isValidImageId = (id) => /^[A-Za-z0-9_-]{1,64}$/.test(id);
const isContentAddressedImageId = (id) => /^[a-f0-9]{24}$/i.test(id);

async function imageMatchesId(blob, id) {
  return !isContentAddressedImageId(id) || (await hashBytes(await blob.arrayBuffer())) === id;
}

export async function saveImageBlob(blob, projectName) {
  const { id, blob: processed } = await processImage(blob);
  const directory = await imageDirectory(projectName, true);
  const existing = await readImageFile(directory, id);
  if (!existing || !(await imageMatchesId(existing, id))) {
    await writeImageFile(directory, id, processed);
  }
  return id;
}

export async function getImageBlob(id, projectName) {
  if (!isValidImageId(id)) throw new Error('invalid image id');
  const directory = await imageDirectory(projectName, false);
  if (!directory) return null;
  const blob = await readImageFile(directory, id);
  if (!blob || !(await imageMatchesId(blob, id))) return null;
  return blob;
}

export async function saveImageFile(id, blob, projectName) {
  if (!isValidImageId(id)) throw new Error('invalid image id');
  if (!(await imageMatchesId(blob, id))) throw new Error('이미지 파일이 손상되었거나 ID와 일치하지 않습니다.');
  const directory = await imageDirectory(projectName, true);
  const existing = await readImageFile(directory, id);
  if (!existing || !(await imageMatchesId(existing, id))) {
    await writeImageFile(directory, id, blob);
  }
}

function imageCacheKey(projectName, id) {
  return JSON.stringify([projectName, id]);
}

export async function getImageUrl(id, projectName) {
  if (!isValidImageId(id)) return null;
  const key = imageCacheKey(projectName, id);
  if (urlCache.has(key)) return urlCache.get(key);
  const blob = await getImageBlob(id, projectName);
  if (!blob) return null;
  const url = URL.createObjectURL(blob);
  urlCache.set(key, url);
  return url;
}

export function releaseProjectImages(projectName) {
  for (const [key, url] of urlCache) {
    if (JSON.parse(key)[0] !== projectName) continue;
    URL.revokeObjectURL(url);
    urlCache.delete(key);
  }
}

function placeholderElement(id, label) {
  const placeholder = document.createElement('span');
  placeholder.className = 'cell-img-missing';
  placeholder.setAttribute('role', 'img');
  placeholder.setAttribute('aria-label', label);
  placeholder.title = `${label} (${id})`;
  placeholder.textContent = `🖼 ${label}`;
  return placeholder;
}

export function hydrateImages(root, missingLabel, projectName) {
  if (!projectName) return;
  root.querySelectorAll('img[data-img-id]:not([data-img-state])').forEach((image) => {
    image.dataset.imgState = 'loading';
    const id = image.dataset.imgId;
    getImageUrl(id, projectName).then((url) => {
      if (!image.isConnected) return;
      if (!url) {
        image.replaceWith(placeholderElement(id, missingLabel));
        return;
      }
      image.onerror = () => image.replaceWith(placeholderElement(id, missingLabel));
      image.dataset.imgState = 'ready';
      image.src = url;
    }).catch((error) => {
      if (!image.isConnected) return;
      console.error(`Could not load image ${id} for project "${projectName}".`, error);
      image.replaceWith(placeholderElement(id, missingLabel));
    });
  });
}
