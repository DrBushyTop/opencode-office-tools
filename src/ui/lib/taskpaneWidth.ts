/// <reference types="@types/office-js" />

const DESKTOP_TASKPANE_WIDTH = 560;
const WEB_TASKPANE_WIDTH = 500;

function requestedTaskpaneWidth() {
  return Office.context.platform === Office.PlatformType.OfficeOnline
    ? WEB_TASKPANE_WIDTH
    : DESKTOP_TASKPANE_WIDTH;
}

export function applyDefaultTaskpaneWidth() {
  const taskpane = Office.extensionLifeCycle?.taskpane;
  if (!taskpane?.setWidth) return;

  try {
    taskpane.setWidth(requestedTaskpaneWidth());
  } catch {
  }
}
