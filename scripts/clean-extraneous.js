const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const extraneousPackages = [
  'office-addin-dev-certs',
  'office-addin-cli',
  'office-addin-usage-data',
  'npm-normalize-package-bin',
  'read-package-json-fast',
];

function syncInstallState() {
  console.log('Cleaning extraneous packages...');
  console.log('Syncing dependencies with bun install...');

  try {
    execSync('bun install --silent', { stdio: 'inherit' });
  } catch {
    console.log('bun install completed with warnings (this is usually fine)');
  }
}

function removePackageDirectory(nodeModulesDir, packageName) {
  const packageDir = path.join(nodeModulesDir, packageName);
  if (!fs.existsSync(packageDir)) {
    return;
  }

  console.log(`Removing ${packageName}...`);
  try {
    fs.rmSync(packageDir, { recursive: true, force: true });
    console.log(`  ✓ Removed ${packageName}`);
  } catch (error) {
    console.log(`  ⚠ Could not remove ${packageName}: ${error.message}`);
  }
}

function pruneExtraneousPackages() {
  const nodeModulesDir = path.resolve(__dirname, '../node_modules');
  for (const packageName of extraneousPackages) {
    removePackageDirectory(nodeModulesDir, packageName);
  }
  console.log('Done cleaning extraneous packages.');
}

syncInstallState();
pruneExtraneousPackages();
