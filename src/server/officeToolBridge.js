class OfficeToolBridge {
  constructor() {
    this.bridgeToken = crypto.randomUUID();
    this.clientSessions = new Map();
    this.executors = new Map();
    this.pending = new Map();
  }

  issueClientSession() {
    const token = crypto.randomUUID();
    this.clientSessions.set(token, { createdAt: Date.now(), lastSeen: Date.now() });
    return token;
  }

  validateClientSession(token) {
    this.cleanup();
    const value = this.clientSessions.get(String(token || ''));
    if (!value) {
      throw new Error('Invalid Office bridge session');
    }
    value.lastSeen = Date.now();
    return true;
  }

  register(host, sessionToken) {
    this.validateClientSession(sessionToken);
    for (const [executorId, executor] of this.executors.entries()) {
      if (executor.host === host) {
        if (executor.sessionToken !== sessionToken && Date.now() - executor.lastSeen < 15000) {
          throw new Error(`An active ${host} Office executor is already registered`);
        }
        this.executors.delete(executorId);
      }
    }

    const executorId = crypto.randomUUID();
    this.executors.set(executorId, { host, sessionToken, lastSeen: Date.now() });
    return { executorId };
  }

  unregister(executorId, sessionToken) {
    const executor = this.getExecutor(executorId, sessionToken);
    this.executors.delete(executorId);

    for (const item of this.pending.values()) {
      if (item.assignedExecutorId === executorId) {
        item.assignedExecutorId = null;
      }
    }

    return executor;
  }

  getExecutor(executorId, sessionToken) {
    this.validateClientSession(sessionToken);
    const executor = this.executors.get(String(executorId || ''));
    if (!executor) {
      throw new Error('Office executor not found');
    }
    if (executor.sessionToken !== sessionToken) {
      throw new Error('Office executor session mismatch');
    }
    if (Date.now() - executor.lastSeen >= 15000 && !this.hasAssignedPendingWork(String(executorId || ''))) {
      this.executors.delete(String(executorId || ''));
      throw new Error('Office executor expired');
    }
    return executor;
  }

  hasExecutor(host) {
    this.cleanup();
    return Array.from(this.executors.values()).some((executor) => executor.host === host);
  }

  heartbeat(executorId, sessionToken) {
    const executor = this.getExecutor(executorId, sessionToken);
    executor.lastSeen = Date.now();
    return executor;
  }

  hosts() {
    this.cleanup();
    return Array.from(this.executors.values()).map((executor) => executor.host);
  }

  async execute(host, toolName, args = {}, bridgeToken) {
    if (bridgeToken !== this.bridgeToken) {
      throw new Error('Invalid Office bridge token');
    }
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
        assignedExecutorId: null,
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

  poll(executorId, sessionToken) {
    const executor = this.heartbeat(executorId, sessionToken);
    const item = Array.from(this.pending.values()).find((pending) => pending.host === executor.host && (!pending.assignedExecutorId || pending.assignedExecutorId === executorId)) || null;
    if (item) {
      item.assignedExecutorId = executorId;
    }
    return item;
  }

  respond(id, executorId, sessionToken, payload) {
    const executor = this.heartbeat(executorId, sessionToken);
    const item = this.pending.get(id);
    if (!item) {
      throw new Error('Office tool request not found');
    }
    if (item.host !== executor.host || item.assignedExecutorId !== executorId) {
      throw new Error('Office tool request does not belong to this executor');
    }

    if (payload && payload.error) {
      item.reject(new Error(payload.error));
      return;
    }

    item.resolve(payload);
  }

  cleanup() {
    const now = Date.now();
    for (const [token, session] of this.clientSessions.entries()) {
      if (now - session.lastSeen >= 60 * 60 * 1000) {
        this.clientSessions.delete(token);
      }
    }

    for (const [executorId, executor] of this.executors.entries()) {
      if (now - executor.lastSeen >= 15000 && !this.hasAssignedPendingWork(executorId)) {
        this.executors.delete(executorId);
        for (const item of this.pending.values()) {
          if (item.assignedExecutorId === executorId) {
            item.assignedExecutorId = null;
          }
        }
      }
    }
  }

  hasAssignedPendingWork(executorId) {
    return Array.from(this.pending.values()).some((item) => item.assignedExecutorId === executorId);
  }
}

module.exports = {
  OfficeToolBridge,
};
