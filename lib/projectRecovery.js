import { parseStudyFile } from './parseSheet.js';
import { projectNameFromFileName, projectNameKey } from './projectStore.js?v=4';

const WORKBOOK_PATTERN = /\.(xlsx|xlsm|csv|txt|tsv)$/i;

function projectFromFile(file, pairs, projectName, fsrsText, errors) {
  let savedProjectId = '';
  let savedPlan = null;
  if (fsrsText !== null) {
    try {
      const progress = JSON.parse(fsrsText);
      if (
        typeof progress.projectId === 'string' &&
        progress.projectId.length > 0 &&
        progress.workbookFileName === file.name
      ) {
        savedProjectId = progress.projectId;
        if (progress.plan && /^\d{4}-\d{2}-\d{2}$/.test(progress.plan.examDate || '') && Number.isFinite(progress.plan.desiredRetention)) {
          savedPlan = {
            examDate: progress.plan.examDate,
            desiredRetention: progress.plan.desiredRetention,
            ...(Number.isInteger(progress.plan.dailyNewCards) ? { dailyNewCards: progress.plan.dailyNewCards } : {}),
            ...(typeof progress.plan.useRecommendedDailyPace === 'boolean' ? { useRecommendedDailyPace: progress.plan.useRecommendedDailyPace } : {}),
            ...(typeof progress.plan.firstPassDeadline === 'string' ? { firstPassDeadline: progress.plan.firstPassDeadline } : {}),
          };
        }
      } else {
        errors.push(`${projectName}/${file.name}: FSRS project information does not match the workbook.`);
      }
    } catch (error) {
      errors.push(`${projectName}/${file.name}: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  return {
    id: savedProjectId || crypto.randomUUID(),
    name: projectName,
    fileName: file.name,
    sourceName: file.name,
    createdAt: file.lastModified || Date.now(),
    pairs,
    title: projectName,
    a4Title: null,
    a4Info: null,
    formatVersion: 12,
    cartridgeColorIndex: 0,
    ...(savedPlan ? { plan: savedPlan } : {}),
  };
}

export async function recoverWorkbookProjects(directory, readProgressFile) {
  const projects = [];
  const errors = [];
  const projectIds = new Set();
  const projectNames = new Set();

  for await (const [folderName, folder] of directory.entries()) {
    if (folder.kind !== 'directory') continue;
    const candidates = [];
    for await (const [fileName, handle] of folder.entries()) {
      if (
        handle.kind === 'file' &&
        WORKBOOK_PATTERN.test(fileName) &&
        projectNameKey(projectNameFromFileName(fileName)) === projectNameKey(folderName)
      ) {
        candidates.push(handle);
      }
    }
    if (!candidates.length) continue;
    if (candidates.length > 1) {
      errors.push(`${folderName}: multiple workbook files match this project folder.`);
      continue;
    }

    const fileHandle = candidates[0];
    let file;
    let pairs;
    try {
      file = await fileHandle.getFile();
      pairs = await parseStudyFile(file);
    } catch (error) {
      errors.push(`${folderName}/${fileHandle.name}: ${error instanceof Error ? error.message : String(error)}`);
      continue;
    }
    if (!pairs.length) continue;

    let fsrsText = null;
    try {
      fsrsText = await readProgressFile(folderName, file.name);
    } catch (error) {
      errors.push(`${folderName}/${file.name}: ${error instanceof Error ? error.message : String(error)}`);
    }
    const project = projectFromFile(file, pairs, folderName, fsrsText, errors);
    const nameKey = projectNameKey(project.name);
    if (projectNames.has(nameKey)) {
      errors.push(`${project.name}: duplicate project folder name; this workbook was not restored.`);
      continue;
    }
    if (projectIds.has(project.id)) {
      errors.push(`${folderName}/${file.name}: duplicate ESPER project identifier; this file was not restored.`);
      continue;
    }
    projectNames.add(nameKey);
    projectIds.add(project.id);
    projects.push(project);
  }

  return { projects, errors };
}
