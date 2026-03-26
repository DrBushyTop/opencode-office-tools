const os = require('os');
const path = require('path');

function bridgeTokenDirectory() {
  return path.join(os.tmpdir(), 'opencode-office-bridge', os.userInfo().username);
}

function bridgeTokenPath(port) {
  return path.join(bridgeTokenDirectory(), `${port}.token`);
}

module.exports = {
  bridgeTokenDirectory,
  bridgeTokenPath,
};
