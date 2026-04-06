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

function expectPath(value, label) {
  const parsed = NonEmptyPathSchema.safeParse(value);
  if (!parsed.success) {
    throw new Error(`Invalid ${label}`);
  }
  return parsed.data;
}

function claimSingleInstance() {
  if (app.requestSingleInstanceLock()) {
    return;
  }

  app.quit();
  process.exit(0);
}

function applyPlatformBootTweaks() {
  if (process.platform === 'darwin') {
    app.dock.hide();
  }
}

function resolveRuntimeRoot() {
  if (app.isPackaged) {
    return expectPath(path.join(process.resourcesPath), 'resources path');
  }
  return expectPath(path.resolve(__dirname, '../..'), 'development resources path');
}

function resolveRuntimePaths(rootDir) {
  const trayIconName = process.platform === 'darwin' ? 'tray-icon.png' : 'tray-icon.ico';
  const trayIconPath = expectPath(path.join(rootDir, 'assets', trayIconName), 'tray icon path');
  const logFilePath = app.isPackaged
    ? expectPath(path.join(app.getPath('userData'), 'logs', 'debug.log'), 'packaged log file path')
    : expectPath(path.join(rootDir, '.opencode', 'debug.log'), 'development log file path');

  return { rootDir, trayIconPath, logFilePath };
}

function buildServerEnvironment(rootDir) {
  return TrayEnvSchema.parse({
    OPENCODE_OFFICE_BASE_PATH: rootDir,
    OPENCODE_OFFICE_DIRECTORY: rootDir,
    OPENCODE_OFFICE_CONFIG_DIR: path.join(rootDir, '.opencode'),
  });
}

function installProcessDiagnostics(logFilePath) {
  process.env.OPENCODE_OFFICE_LOG_FILE = logFilePath;
  logInfo('tray', 'Debug logging enabled', { logFilePath });

  process.on('uncaughtException', (error) => {
    logError('tray', 'Uncaught exception', error);
  });
  process.on('unhandledRejection', (reason) => {
    logError('tray', 'Unhandled rejection', reason);
  });
}

function createServiceAdapter(rootDir) {
  const state = {
    server: null,
    phase: 'stopped',
  };

  function statusSnapshot() {
    return {
      isRunning: state.phase === 'running',
      phase: state.phase,
    };
  }

  async function start() {
    try {
      Object.assign(process.env, buildServerEnvironment(rootDir));
      const serverModulePath = require.resolve('../server-prod.js');
      delete require.cache[serverModulePath];
      const { createServer } = require('../server-prod.js');
      state.phase = 'starting';
      state.server = await createServer();
      state.phase = 'running';
      console.log('Server started successfully');
      logInfo('tray', 'Server started successfully');
    } catch (error) {
      state.server = null;
      state.phase = 'stopped';
      console.error('Failed to start server:', error);
      logError('tray', 'Failed to start server', error);
    }

    return statusSnapshot();
  }

  async function stop() {
    if (state.server) {
      await state.server.close?.();
      state.server = null;
    }
    state.phase = 'stopped';
    console.log('Server stopped');
    logInfo('tray', 'Server stopped');
    return statusSnapshot();
  }

  async function toggle() {
    return state.phase === 'running' ? stop() : start();
  }

  function dispose() {
    state.server?.close?.();
  }

  return {
    statusSnapshot,
    start,
    stop,
    toggle,
    dispose,
  };
}

function describeMenuState(serviceStatus) {
  return {
    statusLabel: serviceStatus.isRunning ? '● Service Running' : '○ Service Stopped',
    toggleLabel: serviceStatus.isRunning ? 'Disable Service' : 'Enable Service',
    tooltip: `OpenCode Office Add-in - ${serviceStatus.isRunning ? 'Running' : 'Stopped'}`,
  };
}

function createTrayIcon(trayIconPath) {
  try {
    let icon = nativeImage.createFromPath(trayIconPath);
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

function createTrayPresenter(paths, service) {
  const tray = new Tray(createTrayIcon(paths.trayIconPath));

  function render() {
    const menuState = describeMenuState(service.statusSnapshot());
    tray.setContextMenu(Menu.buildFromTemplate([
      { label: 'OpenCode Office Add-in', enabled: false },
      { type: 'separator' },
      { label: menuState.statusLabel, enabled: false },
      { label: menuState.toggleLabel, click: async () => {
        await service.toggle();
        render();
      } },
      {
        label: 'Open Debug Log',
        click: () => {
          shell.showItemInFolder(expectPath(paths.logFilePath, 'debug log file path'));
        },
      },
      { type: 'separator' },
      { label: 'Quit', click: () => app.quit() },
    ]));
    tray.setToolTip(menuState.tooltip);
  }

  tray.setToolTip('OpenCode Office Add-in - Starting...');
  if (process.platform === 'win32') {
    tray.on('click', () => tray.popUpContextMenu());
  }

  return { tray, render };
}

function registerAppLifecycle(service) {
  app.on('window-all-closed', (event) => {
    event.preventDefault();
  });

  app.on('before-quit', () => {
    logInfo('tray', 'Application quitting', { logFilePath: process.env.OPENCODE_OFFICE_LOG_FILE || getLogFilePath() });
    service.dispose();
  });
}

async function launchTrayRuntime() {
  const paths = resolveRuntimePaths(resolveRuntimeRoot());
  installProcessDiagnostics(paths.logFilePath);

  const service = createServiceAdapter(paths.rootDir);
  const presenter = createTrayPresenter(paths, service);
  registerAppLifecycle(service);

  presenter.render();
  await service.start();
  presenter.render();
}

claimSingleInstance();
applyPlatformBootTweaks();
app.whenReady().then(() => launchTrayRuntime());
