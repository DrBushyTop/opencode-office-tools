/**
 * Electron System Tray App for OpenCode Office Add-in
 * Runs the server in the background with a tray icon
 */
const { app, Tray, Menu, nativeImage, shell } = require('electron');
const path = require('path');
const { z } = require('zod');
const { logInfo, logError, getLogFilePath } = require('../server/devLogger');

const NonEmptyPathSchema = z.string().trim().min(1);
const TrayEnvSchema = z.object({
  OPENCODE_OFFICE_BASE_PATH: NonEmptyPathSchema,
  OPENCODE_OFFICE_DIRECTORY: NonEmptyPathSchema,
  OPENCODE_OFFICE_CONFIG_DIR: NonEmptyPathSchema,
});

/**
 * @typedef {{
 *   OPENCODE_OFFICE_BASE_PATH: string,
 *   OPENCODE_OFFICE_DIRECTORY: string,
 *   OPENCODE_OFFICE_CONFIG_DIR: string,
 * }} TrayEnv
 */

function parseExternalPath(value, context) {
  const result = NonEmptyPathSchema.safeParse(value);
  if (!result.success) {
    throw new Error(`Invalid ${context}`);
  }
  return result.data;
}

// Prevent multiple instances
const gotTheLock = app.requestSingleInstanceLock();
if (!gotTheLock) {
  app.quit();
  process.exit(0);
}

// Hide from dock on macOS
if (process.platform === 'darwin') {
  app.dock.hide();
}

let tray = null;
let server = null;
let serverRunning = false;
let logFilePath = null;

// Get the resources path (works both in dev and when packaged)
function getResourcesPath() {
  if (app.isPackaged) {
    return parseExternalPath(path.join(process.resourcesPath), 'resources path');
  }
  return parseExternalPath(path.resolve(__dirname, '../..'), 'development resources path');
}

function getIconPath() {
  const resourcesPath = getResourcesPath();
  
  if (process.platform === 'darwin') {
    // macOS: use color PNG for menu bar
    return parseExternalPath(path.join(resourcesPath, 'assets', 'tray-icon.png'), 'macOS tray icon path');
  } else {
    // Windows: use .ico
    return parseExternalPath(path.join(resourcesPath, 'assets', 'tray-icon.ico'), 'Windows tray icon path');
  }
}

function resolveLogFilePath() {
  if (!app.isPackaged) {
    return parseExternalPath(path.join(getResourcesPath(), '.opencode', 'debug.log'), 'development log file path');
  }
  return parseExternalPath(path.join(app.getPath('userData'), 'logs', 'debug.log'), 'packaged log file path');
}

function configureLogging() {
  logFilePath = resolveLogFilePath();
  process.env.OPENCODE_OFFICE_LOG_FILE = logFilePath;
  logInfo('tray', 'Debug logging enabled', { logFilePath });
}

async function startServer() {
  try {
    const root = getResourcesPath();
    /** @type {TrayEnv} */
    const trayEnv = TrayEnvSchema.parse({
      OPENCODE_OFFICE_BASE_PATH: root,
      OPENCODE_OFFICE_DIRECTORY: root,
      OPENCODE_OFFICE_CONFIG_DIR: path.join(root, '.opencode'),
    });
    Object.assign(process.env, trayEnv);
    
    // Clear the module cache to allow re-requiring after stop
    const serverModulePath = require.resolve('../server-prod.js');
    delete require.cache[serverModulePath];
    
    // Import and start the server
    const { createServer } = require('../server-prod.js');
    server = await createServer();
    serverRunning = true;
    console.log('Server started successfully');
    logInfo('tray', 'Server started successfully');
    updateTrayMenu();
  } catch (error) {
    console.error('Failed to start server:', error);
    logError('tray', 'Failed to start server', error);
    serverRunning = false;
    updateTrayMenu();
  }
}

async function stopServer() {
  if (server) {
    await server.close?.();
    server = null;
    serverRunning = false;
    console.log('Server stopped');
    logInfo('tray', 'Server stopped');
    updateTrayMenu();
  }
}

async function toggleServer() {
  if (serverRunning) {
    await stopServer();
  } else {
    await startServer();
  }
}

function updateTrayMenu() {
  if (!tray) return;
  
  const statusLabel = serverRunning ? '● Service Running' : '○ Service Stopped';
  const toggleLabel = serverRunning ? 'Disable Service' : 'Enable Service';
  
  const contextMenu = Menu.buildFromTemplate([
    {
      label: 'OpenCode Office Add-in',
      enabled: false
    },
    { type: 'separator' },
    {
      label: statusLabel,
      enabled: false
    },
    {
      label: toggleLabel,
      click: () => toggleServer()
    },
    {
      label: 'Open Debug Log',
      click: () => {
        if (logFilePath) shell.showItemInFolder(parseExternalPath(logFilePath, 'debug log file path'));
      }
    },
    { type: 'separator' },
    {
      label: 'Quit',
      click: () => {
        app.quit();
      }
    }
  ]);

  tray.setContextMenu(contextMenu);
  tray.setToolTip(`OpenCode Office Add-in - ${serverRunning ? 'Running' : 'Stopped'}`);
}

function createTray() {
  const iconPath = getIconPath();
  let icon;
  
  try {
    icon = nativeImage.createFromPath(iconPath);
    // For macOS menu bar, resize to 16x16 or 22x22
    if (process.platform === 'darwin') {
      icon = icon.resize({ width: 16, height: 16 });
      icon.setTemplateImage(false);  // Use color icon, not monochrome template
    }
  } catch (error) {
    console.error('Failed to load tray icon:', error);
    logError('tray', 'Failed to load tray icon', error);
    // Create a simple fallback icon
    icon = nativeImage.createEmpty();
  }

  tray = new Tray(icon);
  tray.setToolTip('OpenCode Office Add-in - Starting...');

  // Initial menu (will be updated after server starts)
  updateTrayMenu();
  
  // On Windows, clicking the tray icon shows the menu
  if (process.platform === 'win32') {
    tray.on('click', () => {
      tray.popUpContextMenu();
    });
  }
}

app.whenReady().then(async () => {
  configureLogging();
  process.on('uncaughtException', (error) => {
    logError('tray', 'Uncaught exception', error);
  });
  process.on('unhandledRejection', (reason) => {
    logError('tray', 'Unhandled rejection', reason);
  });
  createTray();
  await startServer();
});

app.on('window-all-closed', (e) => {
  // Prevent app from quitting when no windows are open
  e.preventDefault();
});

app.on('before-quit', () => {
  logInfo('tray', 'Application quitting', { logFilePath: process.env.OPENCODE_OFFICE_LOG_FILE || getLogFilePath() });
  server?.close?.();
});
