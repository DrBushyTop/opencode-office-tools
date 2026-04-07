const { spawnSync } = require('child_process');

const targetScript = process.platform === 'darwin'
  ? 'build:installer:mac'
  : process.platform === 'win32'
    ? 'build:installer:win'
    : null;

if (!targetScript) {
  console.error(`Unsupported platform for installer build: ${process.platform}`);
  process.exit(1);
}

const result = spawnSync('bun', ['run', targetScript], {
  stdio: 'inherit',
  shell: process.platform === 'win32',
  env: process.env,
});

process.exit(typeof result.status === 'number' ? result.status : 1);
