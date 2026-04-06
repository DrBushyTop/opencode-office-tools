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

function requirePath(value, label) {
  const parsed = NonEmptyPathSchema.safeParse(value);
  if (!parsed.success) {
    throw new Error(`Invalid ${label}`);
  }
  return parsed.data;
}

function ensureSingleInstance() {
  if (app.requestSingleInstanceLock()) {
    return true;
  }

  app.quit();
  process.exit(0);
}

function hideDockWhenNeeded() {
  if (process.platform === 'darwin') {
    app.dock.hide();
  }
}

function resolveInstallRoot() {
  return app.isPackaged
    ? requirePath(path.join(process.resourcesPath), 'resources path')
    : requirePath(path.resolve(__dirname, '../..'), 'development resources path');
}

function resolveLogFilePath(rootDir) {
  if (!app.isPackaged) {
    return requirePath(path.join(rootDir, '.opencode', 'debug.log'), 'development log file path');
  }

  return requirePath(path.join(app.getPath('userData'), 'logs', 'debug.log'), 'packaged log file path');
}

function resolveTrayIcon(rootDir) {
  const filename = process.platform === 'darwin' ? 'tray-icon.png' : 'tray-icon.ico';
  const platformLabel = process.platform === 'darwin' ? 'macOS tray icon path' : 'Windows tray icon path';
  return requirePath(path.join(rootDir, 'assets', filename), platformLabel);
}

function buildServerEnvironment(rootDir) {
  return TrayEnvSchema.parse({
    OPENCODE_OFFICE_BASE_PATH: rootDir,
    OPENCODE_OFFICE_DIRECTORY: rootDir,
    OPENCODE_OFFICE_CONFIG_DIR: path.join(rootDir, '.opencode'),
  });
}

function createRuntimeState(rootDir) {
  return {
    rootDir,
    tray: null,
    server: null,
    serverRunning: false,
    logFilePath: null,
  };
}

function applyLogging(state) {
  state.logFilePath = resolveLogFilePath(state.rootDir);
  process.env.OPENCODE_OFFICE_LOG_FILE = state.logFilePath;
  logInfo('tray', 'Debug logging enabled', { logFilePath: state.logFilePath });
}

function attachProcessDiagnostics() {
  process.on('uncaughtException', (error) => {
    logError('tray', 'Uncaught exception', error);
  });
  process.on('unhandledRejection', (reason) => {
    logError('tray', 'Unhandled rejection', reason);
  });
}

function describeServiceState(isRunning) {
  return {
    statusLabel: isRunning ? '● Service Running' : '○ Service Stopped',
    toggleLabel: isRunning ? 'Disable Service' : 'Enable Service',
    tooltip: `OpenCode Office Add-in - ${isRunning ? 'Running' : 'Stopped'}`,
  };
}

function refreshTrayUi(state, actions) {
  if (!state.tray) {
    return;
  }

  const presentation = describeServiceState(state.serverRunning);
  state.tray.setContextMenu(Menu.buildFromTemplate([
    { label: 'OpenCode Office Add-in', enabled: false },
    { type: 'separator' },
    { label: presentation.statusLabel, enabled: false },
    { label: presentation.toggleLabel, click: () => actions.toggleServer() },
    {
      label: 'Open Debug Log',
      click: () => {
        if (state.logFilePath) {
          shell.showItemInFolder(requirePath(state.logFilePath, 'debug log file path'));
        }
      },
    },
    { type: 'separator' },
    { label: 'Quit', click: () => app.quit() },
  ]));
  state.tray.setToolTip(presentation.tooltip);
}

function loadTrayIcon(state) {
  try {
    let icon = nativeImage.createFromPath(resolveTrayIcon(state.rootDir));
    if (process.platform === 'darwin') {
      icon = icon.resize({ width: 16, height: 16 });
      icon.setTemplateImage(false);
    }
    return icon;
  } catch (error) {
    console.error('Failed to load tray icon:', error);
    logError('tray', 'Failed to load tray icon', error);
    return nativeImage.createEmpty();
  }
}

function mountTray(state, actions) {
  state.tray = new Tray(loadTrayIcon(state));
  state.tray.setToolTip('OpenCode Office Add-in - Starting...');
  refreshTrayUi(state, actions);

  if (process.platform === 'win32') {
    state.tray.on('click', () => {
      state.tray.popUpContextMenu();
    });
  }
}

async function startServer(state, actions) {
  try {
    Object.assign(process.env, buildServerEnvironment(state.rootDir));

    const serverModulePath = require.resolve('../server-prod.js');
    delete require.cache[serverModulePath];
    const { createServer } = require('../server-prod.js');
    state.server = await createServer();
    state.serverRunning = true;
    console.log('Server started successfully');
    logInfo('tray', 'Server started successfully');
  } catch (error) {
    console.error('Failed to start server:', error);
    logError('tray', 'Failed to start server', error);
    state.server = null;
    state.serverRunning = false;
  }

  refreshTrayUi(state, actions);
}

async function stopServer(state, actions) {
  if (!state.server) {
    return;
  }

  await state.server.close?.();
  state.server = null;
  state.serverRunning = false;
  console.log('Server stopped');
  logInfo('tray', 'Server stopped');
  refreshTrayUi(state, actions);
}

function createActions(state) {
  const actions = {
    startServer: () => startServer(state, actions),
    stopServer: () => stopServer(state, actions),
    toggleServer: async () => {
      if (state.serverRunning) {
        await stopServer(state, actions);
        return;
      }

      await startServer(state, actions);
    },
  };

  return actions;
}

function bindAppLifecycle(state) {
  app.on('window-all-closed', (event) => {
    event.preventDefault();
  });

  app.on('before-quit', () => {
    logInfo('tray', 'Application quitting', { logFilePath: process.env.OPENCODE_OFFICE_LOG_FILE || getLogFilePath() });
    state.server?.close?.();
  });
}

async function bootTrayApplication() {
  const state = createRuntimeState(resolveInstallRoot());
  const actions = createActions(state);

  applyLogging(state);
  attachProcessDiagnostics();
  mountTray(state, actions);
  bindAppLifecycle(state);
  await startServer(state, actions);
}

if (ensureSingleInstance()) {
  hideDockWhenNeeded();
  app.whenReady().then(() => bootTrayApplication());
}
