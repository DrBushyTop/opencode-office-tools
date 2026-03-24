class OfficeToolBridge {
  constructor() {
    this.executors = new Map();
    this.pending = new Map();
  }

  register(host) {
    this.executors.set(host, Date.now());
  }

  unregister(host) {
    this.executors.delete(host);
  }

  hasExecutor(host) {
    const seen = this.executors.get(host);
    return Boolean(seen && Date.now() - seen < 15000);
  }

  heartbeat(host) {
    this.executors.set(host, Date.now());
  }

  async execute(host, toolName, args = {}) {
    if (!this.hasExecutor(host)) {
      throw new Error(`No active ${host} Office executor is available`);
    }

    const id = crypto.randomUUID();
    return new Promise((resolve, reject) => {
      const timeout = setTimeout(() => {
        this.pending.delete(id);
        reject(new Error(`Timed out waiting for ${toolName}`));
      }, 20000);

      this.pending.set(id, {
        id,
        host,
        toolName,
        args,
        createdAt: Date.now(),
        resolve: (value) => {
          clearTimeout(timeout);
          this.pending.delete(id);
          resolve(value);
        },
        reject: (error) => {
          clearTimeout(timeout);
          this.pending.delete(id);
          reject(error instanceof Error ? error : new Error(String(error)));
        },
      });
    });
  }

  poll(host) {
    this.heartbeat(host);
    return Array.from(this.pending.values()).find((item) => item.host === host) || null;
  }

  respond(id, payload) {
    const item = this.pending.get(id);
    if (!item) {
      throw new Error('Office tool request not found');
    }

    if (payload && payload.error) {
      item.reject(new Error(payload.error));
      return;
    }

    item.resolve(payload);
  }
}

module.exports = {
  OfficeToolBridge,
};
