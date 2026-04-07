const path = require('path');
const express = require('express');
const { createHttpRuntime, registerShutdown } = require('./server/httpRuntime');
const { certificateDirectory, readPackagedCredentials } = require('./server/localCertificates');
const { logInfo, logError } = require('./server/devLogger');

const isPkg = typeof process.pkg !== 'undefined';

function getBasePath() {
  if (typeof process.env.OPENCODE_OFFICE_BASE_PATH === 'string' && process.env.OPENCODE_OFFICE_BASE_PATH.trim()) {
    return process.env.OPENCODE_OFFICE_BASE_PATH;
  }
  if (isPkg) {
    return path.dirname(process.execPath);
  }
  return path.resolve(__dirname, '..');
}

function parsePort(value, fallback) {
  const port = Number.parseInt(String(value ?? fallback), 10);
  if (!Number.isFinite(port) || port <= 0) {
    throw new Error(`Invalid port: ${value}`);
  }
  return port;
}

function readCertificates(basePath) {
  try {
    return readPackagedCredentials(basePath);
  } catch (error) {
    logError('server', 'TLS credentials are unavailable', {
      certDir: certificateDirectory(),
      basePath,
      error: error instanceof Error ? error.message : String(error),
    });
    throw error;
  }
}

function attachBuiltFrontend(app, basePath) {
  const distPath = path.join(basePath, 'dist');

  app.use((req, res, next) => {
    const startedAt = Date.now();
    res.on('finish', () => {
      logInfo('static', `${req.method} ${req.originalUrl}`, {
        statusCode: res.statusCode,
        durationMs: Date.now() - startedAt,
      });
    });
    next();
  });

  app.use(express.static(distPath));
  app.get('*path', (req, res) => {
    logInfo('static', 'Serving SPA fallback', { url: req.originalUrl, distPath });
    res.sendFile(path.join(distPath, 'index.html'));
  });
}

async function createServer() {
  const basePath = getBasePath();
  const port = parsePort(process.env.PORT, 52390);
  const bridgePort = parsePort(process.env.BRIDGE_PORT, 52391);
  const runtime = createHttpRuntime({
    port,
    bridgePort,
    ...readCertificates(basePath),
  });

  try {
    attachBuiltFrontend(runtime.app, basePath);
    registerShutdown(runtime.close);
    await runtime.listen({
      onHttpsListening: () => {
        console.log(`OpenCode Office Add-in Server running on https://localhost:${port}`);
        logInfo('server', 'HTTPS server started', { port, basePath, certDir: certificateDirectory() });
      },
      onBridgeListening: () => {
        console.log(`Office bridge available at http://127.0.0.1:${bridgePort}/api`);
        logInfo('server', 'Bridge server started', { port: bridgePort });
      },
    });

    return { httpsServer: runtime.httpsServer, close: runtime.close };
  } catch (error) {
    await runtime.close();
    throw error;
  }
}

module.exports = { createServer };

if (require.main === module) {
  createServer().catch((error) => {
    console.error(error);
    logError('server', 'Failed to create server', error);
  });
}
