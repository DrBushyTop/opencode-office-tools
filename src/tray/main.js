/**
 * Electron System Tray App for OpenCode Office Add-in
 * Runs the server in the background with a tray icon
 */
const { app, Tray, Menu, nativeImage, shell } = require('electron');
const path = require('path');
const { logInfo, logError, getLogFilePath } = require('../server/devLogger');

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
    return path.join(process.resourcesPath);
  }
  return path.resolve(__dirname, '../..');
}

function getIconPath() {
  const resourcesPath = getResourcesPath();
  
  if (process.platform === 'darwin') {
    // macOS: use color PNG for menu bar
    return path.join(resourcesPath, 'assets', 'tray-icon.png');
  } else {
    // Windows: use .ico
    return path.join(resourcesPath, 'assets', 'tray-icon.ico');
  }
}

function resolveLogFilePath() {
  if (!app.isPackaged) {
    return path.join(getResourcesPath(), '.opencode', 'debug.log');
  }
  return path.join(app.getPath('userData'), 'logs', 'debug.log');
}

function configureLogging() {
  logFilePath = resolveLogFilePath();
  process.env.OPENCODE_OFFICE_LOG_FILE = logFilePath;
  logInfo('tray', 'Debug logging enabled', { logFilePath });
}

async function startServer() {
  try {
    const root = getResourcesPath();
    process.env.OPENCODE_OFFICE_BASE_PATH = root;
    process.env.OPENCODE_OFFICE_DIRECTORY = root;
    process.env.OPENCODE_OFFICE_CONFIG_DIR = path.join(root, '.opencode');
    
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
        if (logFilePath) shell.showItemInFolder(logFilePath);
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
