const { app, Tray, Menu, nativeImage, shell } = require('electron');
const path = require('path');
const { logInfo, logError, getLogFilePath } = require('../server/devLogger');

const PRODUCT_NAME = 'OpenCode Office Add-in';

function assertPath(value, label) {
  if (typeof value !== 'string' || value.trim().length === 0) {
    throw new Error(`Invalid ${label}`);
  }
  return value;
}

function exitIfAlreadyRunning() {
  if (!app.requestSingleInstanceLock()) {
    app.quit();
    process.exit(0);
  }
}

function hideDockWhenNeeded() {
  if (process.platform === 'darwin') {
    app.dock.hide();
  }
}

function resolveWorkspaceRoot() {
  if (app.isPackaged) {
    return assertPath(path.join(process.resourcesPath), 'resources path');
  }
  return assertPath(path.resolve(__dirname, '../..'), 'development resources path');
}

function resolveLogFile(rootDir) {
  if (app.isPackaged) {
    return assertPath(path.join(app.getPath('userData'), 'logs', 'debug.log'), 'packaged log file path');
  }
  return assertPath(path.join(rootDir, '.opencode', 'debug.log'), 'development log file path');
}

function resolveTrayState() {
  const rootDir = resolveWorkspaceRoot();
  const iconName = process.platform === 'darwin' ? 'tray-icon.png' : 'tray-icon.ico';
  const userDataDir = app.isPackaged
    ? assertPath(app.getPath('userData'), 'packaged user data path')
    : assertPath(path.join(rootDir, '.opencode'), 'development user data path');

  return {
    rootDir,
    trayIconPath: assertPath(path.join(rootDir, 'assets', iconName), 'tray icon path'),
    logFilePath: resolveLogFile(rootDir),
    userDataDir,
    certDir: assertPath(path.join(userDataDir, 'certs'), 'certificate directory path'),
  };
}

function applyServerEnvironment(paths) {
  Object.assign(process.env, {
    OPENCODE_OFFICE_BASE_PATH: paths.rootDir,
    OPENCODE_OFFICE_DIRECTORY: paths.rootDir,
    OPENCODE_OFFICE_CONFIG_DIR: path.join(paths.rootDir, '.opencode'),
    OPENCODE_OFFICE_USER_DATA_DIR: paths.userDataDir,
    OPENCODE_OFFICE_CERT_DIR: paths.certDir,
  });
}

function installDiagnostics(logFilePath) {
  process.env.OPENCODE_OFFICE_LOG_FILE = logFilePath;
  logInfo('tray', 'Debug logging enabled', { logFilePath });

  process.on('uncaughtException', (error) => {
    logError('tray', 'Uncaught exception', error);
  });
  process.on('unhandledRejection', (reason) => {
    logError('tray', 'Unhandled rejection', reason);
  });
}

function loadTrayIcon(iconPath) {
  try {
    let icon = nativeImage.createFromPath(iconPath);
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

function createRuntimeController(paths) {
  let currentServer = null;
  let phase = 'stopped';
  const subscribers = new Set();

  function snapshot() {
    return {
      isRunning: phase === 'running',
      phase,
    };
  }

  function publish() {
    const state = snapshot();
    subscribers.forEach((subscriber) => subscriber(state));
  }

  async function start() {
    try {
      applyServerEnvironment(paths);
      const serverModulePath = require.resolve('../server-prod.js');
      delete require.cache[serverModulePath];
      const { createServer } = require('../server-prod.js');

      phase = 'starting';
      publish();

      currentServer = await createServer();
      phase = 'running';
      console.log('Server started successfully');
      logInfo('tray', 'Server started successfully');
    } catch (error) {
      currentServer = null;
      phase = 'stopped';
      console.error('Failed to start server:', error);
      logError('tray', 'Failed to start server', error);
    }

    publish();
    return snapshot();
  }

  async function stop() {
    if (currentServer) {
      const serverToClose = currentServer;
      currentServer = null;
      await Promise.resolve(serverToClose.close?.());
    }

    phase = 'stopped';
    console.log('Server stopped');
    logInfo('tray', 'Server stopped');
    publish();
    return snapshot();
  }

  function dispose() {
    void Promise.resolve(currentServer?.close?.());
    currentServer = null;
    phase = 'stopped';
  }

  return {
    dispose,
    start,
    status: snapshot,
    stop,
    subscribe(listener) {
      subscribers.add(listener);
      return () => subscribers.delete(listener);
    },
    toggle() {
      return phase === 'running' ? stop() : start();
    },
  };
}

function menuState(status) {
  return {
    statusLabel: status.isRunning ? '● Service Running' : '○ Service Stopped',
    toggleLabel: status.isRunning ? 'Disable Service' : 'Enable Service',
    tooltip: `${PRODUCT_NAME} - ${status.isRunning ? 'Running' : 'Stopped'}`,
  };
}

function createTrayMenu(paths, controller, state) {
  const view = menuState(state);
  return Menu.buildFromTemplate([
    { label: PRODUCT_NAME, enabled: false },
    { type: 'separator' },
    { label: view.statusLabel, enabled: false },
    {
      label: view.toggleLabel,
      click: async () => {
        await controller.toggle();
      },
    },
    {
      label: 'Open Debug Log',
      click: () => {
        shell.showItemInFolder(assertPath(paths.logFilePath, 'debug log file path'));
      },
    },
    { type: 'separator' },
    { label: 'Quit', click: () => app.quit() },
  ]);
}

function attachTrayUi(paths, controller) {
  const tray = new Tray(loadTrayIcon(paths.trayIconPath));

  function render(nextState = controller.status()) {
    const view = menuState(nextState);
    tray.setContextMenu(createTrayMenu(paths, controller, nextState));
    tray.setToolTip(view.tooltip);
  }

  tray.setToolTip(`${PRODUCT_NAME} - Starting...`);
  if (process.platform === 'win32') {
    tray.on('click', () => tray.popUpContextMenu());
  }

  const detach = controller.subscribe(render);
  render();

  return {
    detach,
    tray,
  };
}

function wireLifecycle(controller) {
  app.on('window-all-closed', (event) => {
    event.preventDefault();
  });

  app.on('before-quit', () => {
    logInfo('tray', 'Application quitting', {
      logFilePath: process.env.OPENCODE_OFFICE_LOG_FILE || getLogFilePath(),
    });
    controller.dispose();
  });
}

async function main() {
  const paths = resolveTrayState();
  installDiagnostics(paths.logFilePath);

  const controller = createRuntimeController(paths);
  attachTrayUi(paths, controller);
  wireLifecycle(controller);

  await controller.start();
}

exitIfAlreadyRunning();
hideDockWhenNeeded();
app.whenReady().then(() => main());
