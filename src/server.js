const path = require('path');
const { createHttpRuntime, registerShutdown } = require('./server/httpRuntime');
const { readDevelopmentCredentials } = require('./server/localCertificates');
const { logInfo, logError } = require('./server/devLogger');

const DEV_PORT = 52390;
const DEV_BRIDGE_PORT = 52391;

function readDevCredentials() {
  return readDevelopmentCredentials(path.resolve(__dirname, '..'));
}

async function attachFrontend(app, httpsServer) {
  const { createServer: createViteServer } = require('vite');
  const vite = await createViteServer({
    server: {
      middlewareMode: true,
      hmr: {
        server: httpsServer,
      },
    },
    appType: 'spa',
    configFile: path.resolve(__dirname, '../vite.config.js'),
  });

  app.use(vite.middlewares);
  return vite;
}

async function createServer() {
  const runtime = createHttpRuntime({
    port: DEV_PORT,
    bridgePort: DEV_BRIDGE_PORT,
    ...readDevCredentials(),
  });
  let vite;

  const close = async () => {
    await Promise.resolve(vite?.close?.());
    await runtime.close();
  };

  try {
    vite = await attachFrontend(runtime.app, runtime.httpsServer);
    registerShutdown(close);
    await runtime.listen({
      onHttpsListening: () => {
        console.log(`Server running on https://localhost:${DEV_PORT}`);
        console.log(`API available at https://localhost:${DEV_PORT}/api`);
        logInfo('server', 'Dev HTTPS server started', { port: DEV_PORT });
      },
      onBridgeListening: () => {
        console.log(`Office bridge available at http://127.0.0.1:${DEV_BRIDGE_PORT}/api`);
        logInfo('server', 'Dev bridge server started', { port: DEV_BRIDGE_PORT });
      },
    });

    return { close };
  } catch (error) {
    await close();
    throw error;
  }
}

module.exports = { createServer };

if (require.main === module) {
  createServer().catch((error) => {
    console.error(error);
    logError('server', 'Failed to create dev server', error);
  });
}
