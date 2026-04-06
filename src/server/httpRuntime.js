const express = require('express');
const https = require('https');
const http = require('http');
const { createApiRouter, createBridgeRouter } = require('./api');
const { OpencodeRuntime } = require('./opencodeRuntime');
const { OfficeToolBridge } = require('./officeToolBridge');
const { writeBridgeToken, removeBridgeToken } = require('./bridgeTokenPath');
const { logInfo, logError } = require('./devLogger');

function onceListening(server, listenArgs, onListening) {
  return new Promise((resolve, reject) => {
    const onError = (error) => {
      server.off('error', onError);
      reject(error);
    };

    server.once('error', onError);
    server.listen(...listenArgs, () => {
      server.off('error', onError);
      onListening?.();
      resolve();
    });
  });
}

function closeServer(server) {
  return new Promise((resolve, reject) => {
    if (!server.listening) {
      resolve();
      return;
    }

    server.close((error) => {
      if (error && error.code !== 'ERR_SERVER_NOT_RUNNING') {
        reject(error);
        return;
      }
      resolve();
    });
  });
}

function registerShutdown(close, scope = 'server') {
  async function handleSignal(signal) {
    logInfo(scope, `Received ${signal}`);
    try {
      await close();
    } catch (error) {
      logError(scope, `Failed shutdown after ${signal}`, error);
    }
    process.exit(0);
  }

  process.once('SIGINT', () => {
    void handleSignal('SIGINT');
  });
  process.once('SIGTERM', () => {
    void handleSignal('SIGTERM');
  });
}

function createHttpRuntime({ port, bridgePort, cert, key }) {
  const app = express();
  const bridgeApp = express();
  const runtime = new OpencodeRuntime();
  const bridge = new OfficeToolBridge();

  writeBridgeToken(bridgePort, bridge.bridgeToken);
  app.use('/api', createApiRouter(runtime, bridge));
  bridgeApp.use('/api', createBridgeRouter(bridge));

  const httpsServer = https.createServer({ cert, key }, app);
  const bridgeServer = http.createServer(bridgeApp);
  let closed = false;

  async function close() {
    if (closed) {
      return;
    }
    closed = true;

    runtime.close();
    await Promise.allSettled([
      closeServer(httpsServer),
      closeServer(bridgeServer),
    ]);
    removeBridgeToken(bridgePort);
  }

  async function listen({ onHttpsListening, onBridgeListening } = {}) {
    await Promise.all([
      onceListening(httpsServer, [port], onHttpsListening),
      onceListening(bridgeServer, [bridgePort, '127.0.0.1'], onBridgeListening),
    ]);
  }

  return {
    app,
    bridgeApp,
    bridgePort,
    close,
    httpsServer,
    listen,
    port,
  };
}

module.exports = {
  createHttpRuntime,
  registerShutdown,
};
