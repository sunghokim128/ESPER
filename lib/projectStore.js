import { normalizeStudyRecords } from './studyRecords.js?v=2';

const DATABASE_NAME = 'esperProjects';
const DATABASE_VERSION = 1;
const PROJECT_STORE = 'projects';
const SETTINGS_STORE = 'settings';
const DIRECTORY_KEY = 'workbookDirectory';
const PROJECT_ORDER_KEY = 'projectOrder';
const STUDY_SETTINGS_KEY = 'studySettings';
const LANGUAGE_KEY = 'language';

function openDatabase() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DATABASE_NAME, DATABASE_VERSION);
    request.onupgradeneeded = () => {
      const database = request.result;
      if (!database.objectStoreNames.contains(PROJECT_STORE)) {
        database.createObjectStore(PROJECT_STORE, { keyPath: 'id' });
      }
      if (!database.objectStoreNames.contains(SETTINGS_STORE)) {
        database.createObjectStore(SETTINGS_STORE, { keyPath: 'key' });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error('프로젝트 저장소를 열지 못했어요.'));
  });
}

async function withStore(storeName, mode, operation) {
  const database = await openDatabase();
  try {
    return await new Promise((resolve, reject) => {
      const transaction = database.transaction(storeName, mode);
      let result;
      const request = operation(transaction.objectStore(storeName));
      request.onsuccess = () => {
        result = request.result;
      };
      request.onerror = () => reject(request.error || new Error('프로젝트 정보를 처리하지 못했어요.'));
      transaction.oncomplete = () => resolve(result);
      transaction.onerror = () => reject(transaction.error || new Error('프로젝트 정보를 처리하지 못했어요.'));
      transaction.onabort = () => reject(transaction.error || new Error('프로젝트 정보 처리가 취소됐어요.'));
    });
  } finally {
    database.close();
  }
}

export async function saveWorkbookDirectory(handle) {
  await withStore(SETTINGS_STORE, 'readwrite', (store) =>
    store.put({ key: DIRECTORY_KEY, handle })
  );
}

export async function getStoredWorkbookDirectory() {
  const record = await withStore(SETTINGS_STORE, 'readonly', (store) =>
    store.get(DIRECTORY_KEY)
  );
  return record?.handle || null;
}

export async function getProjectOrder() {
  const record = await withStore(SETTINGS_STORE, 'readonly', (store) =>
    store.get(PROJECT_ORDER_KEY)
  );
  return Array.isArray(record?.value) ? record.value : [];
}

export async function saveProjectOrder(projectIds) {
  await withStore(SETTINGS_STORE, 'readwrite', (store) =>
    store.put({ key: PROJECT_ORDER_KEY, value: projectIds })
  );
}

const LIBRARY_FILE_NAME = 'esper-library.json';

export async function readLibraryFile() {
  const directory = await getWorkbookDirectory();
  if (!directory) return null;
  try {
    const handle = await directory.getFileHandle(LIBRARY_FILE_NAME);
    const data = JSON.parse(await (await handle.getFile()).text());
    return data && typeof data === 'object' ? data : null;
  } catch {
    return null;
  }
}

export async function writeLibraryFile(data) {
  const directory = await getWorkbookDirectory();
  if (!directory) return false;
  const handle = await directory.getFileHandle(LIBRARY_FILE_NAME, { create: true });
  await writeFileHandle(handle, new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' }));
  return true;
}

export async function getStudySettings() {
  const record = await withStore(SETTINGS_STORE, 'readonly', (store) =>
    store.get(STUDY_SETTINGS_KEY)
  );
  return record?.value || null;
}

export async function saveStudySettings(settings) {
  await withStore(SETTINGS_STORE, 'readwrite', (store) =>
    store.put({ key: STUDY_SETTINGS_KEY, value: settings })
  );
}

export async function getLanguage() {
  const record = await withStore(SETTINGS_STORE, 'readonly', (store) =>
    store.get(LANGUAGE_KEY)
  );
  return record?.value === 'ko' || record?.value === 'en' ? record.value : null;
}

export async function saveLanguage(language) {
  await withStore(SETTINGS_STORE, 'readwrite', (store) =>
    store.put({ key: LANGUAGE_KEY, value: language })
  );
}

export async function getWorkbookDirectory(requestPermission = false) {
  const handle = await getStoredWorkbookDirectory();
  if (!handle) return null;

  const options = { mode: 'readwrite' };
  let permission = await handle.queryPermission(options);
  if (permission !== 'granted' && requestPermission) {
    permission = await handle.requestPermission(options);
  }
  return permission === 'granted' ? handle : null;
}

export function projectNameKey(name) {
  return String(name).normalize('NFC').toLowerCase();
}

export function projectNameFromFileName(fileName) {
  return String(fileName).replace(/\.(xlsx|xlsm|csv|txt|tsv)$/i, '');
}

export function projectNameExists(existingNames, projectName) {
  const wanted = projectNameKey(projectName);
  return existingNames.some((name) => projectNameKey(name) === wanted);
}

function assertValidProjectName(name) {
  if (
    typeof name !== 'string' ||
    !name.trim() ||
    name === '.' ||
    name === '..' ||
    /[. ]$/.test(name) ||
    /^(con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/i.test(name) ||
    /[\\/:*?"<>|\u0000-\u001f]/.test(name)
  ) {
    throw new Error('프로젝트 이름으로 저장 폴더를 만들 수 없습니다. 파일 이름을 확인해 주세요.');
  }
}

async function getProjectDirectory(projectName) {
  assertValidProjectName(projectName);
  const root = await getWorkbookDirectory();
  if (!root) throw new Error('저장 폴더 접근 권한이 없어요. 저장 위치를 다시 선택해 주세요.');
  try {
    return await root.getDirectoryHandle(projectName);
  } catch (error) {
    if (error instanceof DOMException && error.name === 'NotFoundError') return null;
    throw error;
  }
}

export async function createProjectDirectory(projectName) {
  assertValidProjectName(projectName);
  const root = await getWorkbookDirectory();
  if (!root) throw new Error('저장 폴더 접근 권한이 없어요. 저장 위치를 다시 선택해 주세요.');
  const existingNames = [];
  for await (const [name] of root.entries()) {
    existingNames.push(name);
  }
  if (projectNameExists(existingNames, projectName)) {
    throw new Error(`"${projectName}" 이름의 파일 또는 프로젝트가 이미 있어요. 중복되지 않는 파일을 선택해 주세요.`);
  }
  return root.getDirectoryHandle(projectName, { create: true });
}

export async function removeProjectDirectory(projectName) {
  const root = await getWorkbookDirectory();
  if (!root) throw new Error('저장 폴더 접근 권한이 없어요. 저장 위치를 다시 선택해 주세요.');
  try {
    await root.removeEntry(projectName, { recursive: true });
  } catch (error) {
    if (error instanceof DOMException && error.name === 'NotFoundError') return;
    throw error;
  }
}

function studyRecordFileName(projectName) {
  return `${projectName} study record.json`;
}

function fsrsDataFileName(workbookFileName) {
  return `${String(workbookFileName || '').replace(/\.[^.]+$/, '')}.json`;
}

async function writeFileHandle(handle, blob) {
  const writable = await handle.createWritable();
  try {
    await writable.write(blob);
    await writable.close();
  } catch (error) {
    try {
      await writable.abort();
    } catch (abortError) {
      throw new AggregateError([error, abortError], '저장 도중 오류가 발생했고 임시 파일도 취소하지 못했습니다.');
    }
    throw error;
  }
}

export async function readFsrsDataFile(projectName, workbookFileName) {
  const directory = await getProjectDirectory(projectName);
  if (!directory) return null;
  try {
    const handle = await directory.getFileHandle(fsrsDataFileName(workbookFileName));
    return await (await handle.getFile()).text();
  } catch (error) {
    if (error instanceof DOMException && error.name === 'NotFoundError') return null;
    throw error;
  }
}

export async function writeFsrsDataFile(projectName, workbookFileName, contents) {
  const directory = await getProjectDirectory(projectName);
  if (!directory) throw new DOMException('The project folder was not found.', 'NotFoundError');
  const handle = await directory.getFileHandle(fsrsDataFileName(workbookFileName), { create: true });
  await writeFileHandle(handle, new Blob([contents], { type: 'application/json' }));
}

export async function readStudyRecordFile(projectName) {
  const directory = await getProjectDirectory(projectName);
  if (!directory) return null;
  try {
    const handle = await directory.getFileHandle(studyRecordFileName(projectName));
    return normalizeStudyRecords(JSON.parse(await (await handle.getFile()).text()));
  } catch (error) {
    if (error instanceof DOMException && error.name === 'NotFoundError') return null;
    throw error;
  }
}

export async function writeStudyRecordFile(projectName, records) {
  const directory = await getProjectDirectory(projectName);
  if (!directory) throw new DOMException('The project folder was not found.', 'NotFoundError');
  const handle = await directory.getFileHandle(studyRecordFileName(projectName), { create: true });
  await writeFileHandle(handle, new Blob([JSON.stringify(records, null, 2)], { type: 'application/json' }));
}

export async function saveProject(project) {
  await withStore(PROJECT_STORE, 'readwrite', (store) => store.put(project));
}

export async function getProject(id) {
  return withStore(PROJECT_STORE, 'readonly', (store) => store.get(id));
}

export async function listProjects() {
  const projects = await withStore(PROJECT_STORE, 'readonly', (store) => store.getAll());
  return projects.sort((a, b) => b.createdAt - a.createdAt);
}

export async function deleteProjectRecord(id) {
  await withStore(PROJECT_STORE, 'readwrite', (store) => store.delete(id));
}

export async function writeProjectWorkbook(projectName, fileName, blob) {
  const directory = await getProjectDirectory(projectName);
  if (!directory) throw new DOMException('The project folder was not found.', 'NotFoundError');
  const handle = await directory.getFileHandle(fileName, { create: true });
  await writeFileHandle(handle, blob);
}

export async function readProjectWorkbook(projectName, fileName) {
  const directory = await getProjectDirectory(projectName);
  if (!directory) throw new DOMException('The project folder was not found.', 'NotFoundError');
  return (await directory.getFileHandle(fileName)).getFile();
}
