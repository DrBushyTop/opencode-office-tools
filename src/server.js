const express = require('express');
const https = require('https');
const http = require('http');
const { createServer: createViteServer } = require('vite');
const path = require('path');
const fs = require('fs');
const { createApiRouter } = require('./server/api');
const { OpencodeRuntime } = require('./server/opencodeRuntime');
const { OfficeToolBridge } = require('./server/officeToolBridge');

async function createServer() {
  const app = express();
  const runtime = new OpencodeRuntime();
  const bridge = new OfficeToolBridge();
  const apiRouter = createApiRouter(runtime, bridge);
  app.use('/api', apiRouter);

  // ========== Vite Dev Server (Frontend) ==========
  
  // Create HTTPS server first
  const certPath = path.resolve(__dirname, '../certs/localhost.pem');
  const keyPath = path.resolve(__dirname, '../certs/localhost-key.pem');
  
  const httpsConfig = {
    cert: fs.readFileSync(certPath),
    key: fs.readFileSync(keyPath),
  };
  
  const PORT = 52390;
  const BRIDGE_PORT = 52391;
  const httpsServer = https.createServer(httpsConfig, app);
  const bridgeServer = http.createServer(app);
  
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
  });
  bridgeServer.listen(BRIDGE_PORT, '127.0.0.1', () => {
    console.log(`Office bridge available at http://127.0.0.1:${BRIDGE_PORT}/api`);
  });

  const close = () => {
    runtime.close();
    httpsServer.close();
    bridgeServer.close();
  };

  process.once('SIGINT', () => {
    close();
    process.exit(0);
  });
  process.once('SIGTERM', () => {
    close();
    process.exit(0);
  });

  return { close };
}

createServer().catch(console.error);

