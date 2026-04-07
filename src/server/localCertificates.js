const fs = require('fs');
const os = require('os');
const path = require('path');

const PRODUCT_NAME = 'OpenCode Office Add-in';
const WINDOWS_PFX_PASSPHRASE = 'OpenCodeOfficeLocalCert';

function defaultUserDataDirectory() {
  if (process.platform === 'darwin') {
    return path.join(os.homedir(), 'Library', 'Application Support', PRODUCT_NAME);
  }

  if (process.platform === 'win32') {
    const appData = process.env.APPDATA || path.join(os.homedir(), 'AppData', 'Roaming');
    return path.join(appData, PRODUCT_NAME);
  }

  return path.join(os.homedir(), '.config', PRODUCT_NAME);
}

function userDataDirectory() {
  return process.env.OPENCODE_OFFICE_USER_DATA_DIR
    ? path.resolve(process.env.OPENCODE_OFFICE_USER_DATA_DIR)
    : defaultUserDataDirectory();
}

function certificateDirectory() {
  return process.env.OPENCODE_OFFICE_CERT_DIR
    ? path.resolve(process.env.OPENCODE_OFFICE_CERT_DIR)
    : path.join(userDataDirectory(), 'certs');
}

function readRequiredFile(filePath, label) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`Missing ${label} at ${filePath}`);
  }

  return fs.readFileSync(filePath);
}

function readPemCredentials(directory, label) {
  return {
    cert: readRequiredFile(path.join(directory, 'localhost.pem'), `${label} certificate`),
    key: readRequiredFile(path.join(directory, 'localhost-key.pem'), `${label} key`),
  };
}

function readDevelopmentCredentials(basePath) {
  return readPemCredentials(path.join(basePath, 'certs'), 'development');
}

function readPackagedCredentials(basePath) {
  const installedCertDir = certificateDirectory();
  const pfxPath = path.join(installedCertDir, 'localhost.pfx');

  if (fs.existsSync(pfxPath)) {
    return {
      pfx: fs.readFileSync(pfxPath),
      passphrase: WINDOWS_PFX_PASSPHRASE,
      source: pfxPath,
    };
  }

  const installedCertPath = path.join(installedCertDir, 'localhost.pem');
  const installedKeyPath = path.join(installedCertDir, 'localhost-key.pem');
  if (fs.existsSync(installedCertPath) && fs.existsSync(installedKeyPath)) {
    return {
      ...readPemCredentials(installedCertDir, 'installed'),
      source: installedCertDir,
    };
  }

  const bundledCertDir = path.join(basePath, 'certs');
  return {
    ...readPemCredentials(bundledCertDir, 'bundled'),
    source: bundledCertDir,
  };
}

module.exports = {
  PRODUCT_NAME,
  WINDOWS_PFX_PASSPHRASE,
  certificateDirectory,
  defaultUserDataDirectory,
  readDevelopmentCredentials,
  readPackagedCredentials,
  userDataDirectory,
};
