const fs = require('fs');
const os = require('os');
const path = require('path');
const { randomUUID } = require('crypto');

function isPosix() {
  return process.platform !== 'win32' && typeof process.getuid === 'function';
}

function currentUserSuffix() {
  return isPosix() ? String(process.getuid()) : os.userInfo().username;
}

function assertOwnedByCurrentUser(stats, filePath) {
  if (!isPosix()) return;
  if (stats.uid !== process.getuid()) {
    throw new Error(`Bridge token path is not owned by the current user: ${filePath}`);
  }
}

function assertSafePath(filePath, expectedType) {
  const stats = fs.lstatSync(filePath);
  if (stats.isSymbolicLink()) {
    throw new Error(`Refusing to use symbolic link for Office bridge token ${expectedType}: ${filePath}`);
  }
  if (expectedType === 'directory' && !stats.isDirectory()) {
    throw new Error(`Office bridge token directory is not a directory: ${filePath}`);
  }
  if (expectedType === 'file' && !stats.isFile()) {
    throw new Error(`Office bridge token path is not a file: ${filePath}`);
  }
  assertOwnedByCurrentUser(stats, filePath);
  return stats;
}

function bridgeTokenDirectory() {
  return path.join(os.tmpdir(), `opencode-office-bridge-${currentUserSuffix()}`);
}

function bridgeTokenPath(port) {
  return path.join(bridgeTokenDirectory(), `${port}.token`);
}

function ensureBridgeTokenDirectory() {
  const directory = bridgeTokenDirectory();
  if (!fs.existsSync(directory)) {
    fs.mkdirSync(directory, { recursive: true, mode: 0o700 });
  }
  assertSafePath(directory, 'directory');
  if (isPosix()) {
    fs.chmodSync(directory, 0o700);
  }
  return directory;
}

function assertSafeTokenFile(filePath) {
  const stats = assertSafePath(filePath, 'file');
  if (isPosix() && (stats.mode & 0o077) !== 0) {
    throw new Error(`Office bridge token file permissions are too broad: ${filePath}`);
  }
  return stats;
}

function writeBridgeToken(port, token) {
  ensureBridgeTokenDirectory();
  const filePath = bridgeTokenPath(port);
  const tempPath = `${filePath}.${process.pid}.${randomUUID()}.tmp`;
  fs.writeFileSync(tempPath, String(token), { encoding: 'utf8', flag: 'wx', mode: 0o600 });

  try {
    if (isPosix()) {
      fs.chmodSync(tempPath, 0o600);
    }
    assertSafePath(tempPath, 'file');
    if (fs.existsSync(filePath)) {
      assertSafeTokenFile(filePath);
      fs.unlinkSync(filePath);
    }
    fs.renameSync(tempPath, filePath);
    if (isPosix()) {
      fs.chmodSync(filePath, 0o600);
    }
    return filePath;
  } catch (error) {
    if (fs.existsSync(tempPath)) {
      fs.unlinkSync(tempPath);
    }
    throw error;
  }
}

function readBridgeToken(port) {
  const filePath = bridgeTokenPath(port);
  assertSafeTokenFile(filePath);
  return fs.readFileSync(filePath, 'utf8').trim();
}

function removeBridgeToken(port) {
  const filePath = bridgeTokenPath(port);
  if (!fs.existsSync(filePath)) return;
  assertSafeTokenFile(filePath);
  fs.unlinkSync(filePath);
}

module.exports = {
  ensureBridgeTokenDirectory,
  bridgeTokenDirectory,
  bridgeTokenPath,
  writeBridgeToken,
  readBridgeToken,
  removeBridgeToken,
};
