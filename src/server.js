const express = require('express');
const https = require('https');
const http = require('http');
const { createServer: createViteServer } = require('vite');
const path = require('path');
const fs = require('fs');
const { z } = require('zod');
const { createApiRouter, createBridgeRouter } = require('./server/api');
const { OpencodeRuntime } = require('./server/opencodeRuntime');
const { OfficeToolBridge } = require('./server/officeToolBridge');
const { writeBridgeToken, removeBridgeToken } = require('./server/bridgeTokenPath');
const { logInfo, logError } = require('./server/devLogger');

const devServerConfigSchema = z.object({
  PORT: z.number(),
  BRIDGE_PORT: z.number(),
  cert: z.instanceof(Buffer),
  key: z.instanceof(Buffer),
});

async function createServer() {
  const certPath = path.resolve(__dirname, '../certs/localhost.pem');
  const keyPath = path.resolve(__dirname, '../certs/localhost-key.pem');
  const { PORT, BRIDGE_PORT, cert, key } = devServerConfigSchema.parse({
    PORT: 52390,
    BRIDGE_PORT: 52391,
    cert: fs.readFileSync(certPath),
    key: fs.readFileSync(keyPath),
  });
  const app = express();
  const bridgeApp = express();
  const runtime = new OpencodeRuntime();
  const bridge = new OfficeToolBridge();
  writeBridgeToken(BRIDGE_PORT, bridge.bridgeToken);
  const apiRouter = createApiRouter(runtime, bridge);
  app.use('/api', apiRouter);
  bridgeApp.use('/api', createBridgeRouter(bridge));

  // ========== Vite Dev Server (Frontend) ==========
  
  // Create HTTPS server first
  const httpsConfig = { cert, key };

  const httpsServer = https.createServer(httpsConfig, app);
  const bridgeServer = http.createServer(bridgeApp);
  
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

  // Use vite's connect instance as middleware
  app.use(vite.middlewares);

  httpsServer.listen(PORT, () => {
    console.log(`Server running on https://localhost:${PORT}`);
    console.log(`API available at https://localhost:${PORT}/api`);
    logInfo('server', 'Dev HTTPS server started', { port: PORT });
  });
  bridgeServer.listen(BRIDGE_PORT, '127.0.0.1', () => {
    console.log(`Office bridge available at http://127.0.0.1:${BRIDGE_PORT}/api`);
    logInfo('server', 'Dev bridge server started', { port: BRIDGE_PORT });
  });

  const close = () => {
    runtime.close();
    httpsServer.close();
    bridgeServer.close();
    removeBridgeToken(BRIDGE_PORT);
  };

  process.once('SIGINT', () => {
    logInfo('server', 'Received SIGINT');
    close();
    process.exit(0);
  });
  process.once('SIGTERM', () => {
    logInfo('server', 'Received SIGTERM');
    close();
    process.exit(0);
  });

  return { close };
}

createServer().catch((error) => {
  console.error(error);
  logError('server', 'Failed to create dev server', error);
});
