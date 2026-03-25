const { spawn } = require('child_process');
const path = require('path');
const net = require('net');

const DEFAULT_HOST = '127.0.0.1';
const DEFAULT_PORT = 4096;
const OFFICE_ROOT = path.resolve(__dirname, '..', '..');

function officeConfigDirectory() {
  return process.env.OPENCODE_OFFICE_CONFIG_DIR
    ? path.resolve(process.env.OPENCODE_OFFICE_CONFIG_DIR)
    : path.join(officeDirectory(), '.opencode');
}

function trimSlash(value) {
  return String(value || '').replace(/\/+$/, '');
}

function officeDirectory() {
  return process.env.OPENCODE_OFFICE_DIRECTORY
    ? path.resolve(process.env.OPENCODE_OFFICE_DIRECTORY)
    : OFFICE_ROOT;
}

function configuredBaseUrl() {
  const value = process.env.OPENCODE_OFFICE_RUNTIME_URL || process.env.OPENCODE_RUNTIME_URL;
  return value ? trimSlash(value) : '';
}

function readJson(stream) {
  if (!stream) return;
  stream.setEncoding('utf8');
}

function isReachablePort(port) {
  return new Promise((resolve) => {
    const server = net.createServer();
    server.once('error', () => resolve(false));
    server.listen(port, DEFAULT_HOST, () => {
      server.close(() => resolve(true));
    });
  });
}

async function getFreePort(start = DEFAULT_PORT) {
  for (let port = start; port < start + 50; port += 1) {
    if (await isReachablePort(port)) {
      return port;
    }
  }
  throw new Error('Could not find an available OpenCode port');
}

class OpencodeRuntime {
  constructor() {
    this.runtime = null;
    this.starting = null;
    this.cleanup = null;
  }

  directory() {
    return officeDirectory();
  }

  headers(extra = {}) {
    return {
      ...extra,
      'x-opencode-directory': encodeURIComponent(this.directory()),
    };
  }

  async ensureRuntime() {
    if (this.runtime) return this.runtime;
    if (this.starting) return this.starting;

    this.starting = (async () => {
      const attached = await this.tryAttach();
      if (attached) {
        this.runtime = attached;
        return attached;
      }

      const spawned = await this.spawn();
      this.runtime = spawned;
      return spawned;
    })();

    try {
      return await this.starting;
    } finally {
      this.starting = null;
    }
  }

  async tryAttach() {
    const baseUrl = configuredBaseUrl();
    if (!baseUrl) return null;

    try {
      await this.request('/provider', { baseUrl });
      return { mode: 'attached', baseUrl, child: null };
    } catch {
      return null;
    }
  }

  async spawn() {
    const port = process.env.OPENCODE_OFFICE_PORT
      ? Number(process.env.OPENCODE_OFFICE_PORT)
      : await getFreePort(DEFAULT_PORT);
    const baseUrl = `http://${DEFAULT_HOST}:${port}`;

    const child = spawn('opencode', [`serve`, `--hostname=${DEFAULT_HOST}`, `--port=${port}`], {
      env: {
        ...process.env,
        NODE_TLS_REJECT_UNAUTHORIZED: '0',
        OPENCODE_CONFIG_DIR: officeConfigDirectory(),
      },
      stdio: ['ignore', 'pipe', 'pipe'],
    });

    readJson(child.stdout);
    readJson(child.stderr);

    let output = '';
    child.stdout.on('data', (chunk) => {
      output += chunk;
      const text = chunk.toString().trim();
      if (text) console.log(`[opencode] ${text}`);
    });
    child.stderr.on('data', (chunk) => {
      output += chunk;
      const text = chunk.toString().trim();
      if (text) console.error(`[opencode] ${text}`);
    });

    const ready = await new Promise((resolve, reject) => {
      const timeout = setTimeout(() => {
        reject(new Error('Timed out waiting for OpenCode runtime to start'));
      }, 15000);

      const stop = () => clearTimeout(timeout);
      const fail = (error) => {
        stop();
        reject(error instanceof Error ? error : new Error(String(error)));
      };

      child.once('error', fail);
      child.once('exit', (code) => {
        fail(new Error(`OpenCode runtime exited with code ${code}${output ? `\n${output}` : ''}`));
      });

      const check = async () => {
        try {
          await this.request('/provider', { baseUrl });
          stop();
          resolve({ mode: 'spawned', baseUrl, child });
        } catch {
          setTimeout(check, 250);
        }
      };

      check();
    });

    const shutdown = () => {
      if (!child.killed) child.kill('SIGTERM');
    };
    this.cleanup = shutdown;

    return ready;
  }

  close() {
    this.cleanup?.();
    this.cleanup = null;
  }

  async request(url, options = {}) {
    const runtime = options.baseUrl ? { baseUrl: options.baseUrl } : await this.ensureRuntime();
    const response = await fetch(`${runtime.baseUrl}${url}`, {
      method: options.method || 'GET',
      headers: this.headers(options.headers),
      body: options.body ? JSON.stringify(options.body) : undefined,
    });

    if (!response.ok) {
      const text = await response.text();
      throw new Error(text || `OpenCode request failed: ${response.status}`);
    }

    if (options.raw) {
      return response;
    }

    return response.json();
  }

  async listModels() {
    const providers = await this.request('/provider');
    const items = [];

    for (const provider of providers.all || []) {
      for (const model of Object.values(provider.models || {})) {
        const value = `${provider.id}/${model.id}`;
        items.push({
          key: value,
          label: `${provider.name || provider.id} / ${model.name || model.id}`,
          providerID: provider.id,
          modelID: model.id,
        });
      }
    }

    return items;
  }

  async status() {
    const runtime = await this.ensureRuntime();
    let models = [];

    try {
      models = await this.listModels();
    } catch (error) {
      console.warn('Failed to list OpenCode models:', error.message);
    }

    return {
      mode: runtime.mode,
      baseUrl: runtime.baseUrl,
      directory: this.directory(),
      configDirectory: officeConfigDirectory(),
      models,
    };
  }
}

module.exports = {
  OpencodeRuntime,
  officeDirectory,
  officeConfigDirectory,
};
