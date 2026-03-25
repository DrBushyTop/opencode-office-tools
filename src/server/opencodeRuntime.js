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

function parseModelKey(value) {
  const [providerID, ...rest] = String(value || '').split('/');
  const modelID = rest.join('/');
  if (!providerID || !modelID) return null;
  return { providerID, modelID };
}

function toModelInfo(provider, model) {
  const value = `${provider.id}/${model.id}`;
  return {
    key: value,
    label: `${provider.name || provider.id} / ${model.name || model.id}`,
    providerID: provider.id,
    modelID: model.id,
  };
}

function configuredModels(configProviders, config) {
  const providerItems = Array.isArray(configProviders?.providers) ? configProviders.providers : [];
  const defaults = configProviders?.default && typeof configProviders.default === 'object'
    ? Object.values(configProviders.default)
    : [];
  const configured = new Set([config?.model, ...defaults].filter(Boolean));

  if (configured.size === 0) return [];

  const items = [];
  const seen = new Set();

  for (const key of configured) {
    const parsed = parseModelKey(key);
    if (!parsed) continue;
    const provider = providerItems.find((item) => item.id === parsed.providerID);
    const model = provider?.models?.[parsed.modelID];
    const modelInfo = provider && model
      ? toModelInfo(provider, model)
      : {
          key,
          label: key,
          providerID: parsed.providerID,
          modelID: parsed.modelID,
        };

    if (seen.has(modelInfo.key)) continue;
    seen.add(modelInfo.key);
    items.push(modelInfo);
  }

  return items;
}

async function readResponseBody(response) {
  if (response.status === 204 || response.status === 205 || response.status === 304) {
    return null;
  }

  const text = await response.text();
  if (!text) return null;

  const contentType = response.headers.get('content-type') || '';
  if (contentType.includes('application/json')) {
    return JSON.parse(text);
  }

  try {
    return JSON.parse(text);
  } catch {
    return text;
  }
}

async function sdk() {
  return import('@opencode-ai/sdk');
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
    const dir = officeConfigDirectory();
    const env = process.env.OPENCODE_CONFIG_DIR;
    process.env.OPENCODE_CONFIG_DIR = dir;

    try {
      const mod = await sdk();
      const server = await mod.createOpencodeServer({
        hostname: DEFAULT_HOST,
        port,
        timeout: 15000,
      });
      const baseUrl = server.url.replace(/\/+$/, '');
      this.cleanup = () => server.close();
      return { mode: 'spawned', baseUrl, server };
    } finally {
      if (env) process.env.OPENCODE_CONFIG_DIR = env;
      else delete process.env.OPENCODE_CONFIG_DIR;
    }
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

    return readResponseBody(response);
  }

  async listModels() {
    const [configProviders, config] = await Promise.all([
      this.request('/config/providers'),
      this.request('/config').catch(() => null),
    ]);
    const narrowed = configuredModels(configProviders, config);
    if (narrowed.length > 0) return narrowed;

    const providers = await this.request('/provider');
    const items = [];

    for (const provider of providers.all || []) {
      for (const model of Object.values(provider.models || {})) {
        items.push(toModelInfo(provider, model));
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
  configuredModels,
  officeDirectory,
  officeConfigDirectory,
  parseModelKey,
  readResponseBody,
};
