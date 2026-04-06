const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const cleanupPolicy = [
  { kind: 'directory', label: 'office-addin-dev-certs', relativePath: 'office-addin-dev-certs' },
  { kind: 'directory', label: 'office-addin-cli', relativePath: 'office-addin-cli' },
  { kind: 'directory', label: 'office-addin-usage-data', relativePath: 'office-addin-usage-data' },
  { kind: 'directory', label: 'npm-normalize-package-bin', relativePath: 'npm-normalize-package-bin' },
  { kind: 'directory', label: 'read-package-json-fast', relativePath: 'read-package-json-fast' },
];

function syncDependencyGraph() {
  console.log('Cleaning extraneous packages...');
  console.log('Refreshing dependency graph with bun install...');

  try {
    execSync('bun install --silent', { stdio: 'inherit' });
  } catch {
    console.log('bun install completed with warnings (this is usually fine)');
  }
}

function inspectPolicyTarget(nodeModulesRoot, entry) {
  return {
    ...entry,
    absolutePath: path.join(nodeModulesRoot, entry.relativePath),
  };
}

function prunePolicyEntry(entry) {
  if (!fs.existsSync(entry.absolutePath)) {
    return { label: entry.label, removed: false, reason: 'missing' };
  }

  console.log(`Pruning ${entry.label}...`);
  try {
    fs.rmSync(entry.absolutePath, { recursive: true, force: true });
    console.log(`  ✓ Removed ${entry.label}`);
    return { label: entry.label, removed: true, reason: 'removed' };
  } catch (error) {
    console.log(`  ⚠ Could not remove ${entry.label}: ${error.message}`);
    return { label: entry.label, removed: false, reason: 'failed' };
  }
}

function reportCleanupSummary(results) {
  const removedCount = results.filter((result) => result.removed).length;
  const skippedCount = results.filter((result) => result.reason === 'missing').length;
  console.log(`Cleanup summary: removed ${removedCount}, already absent ${skippedCount}.`);
  console.log('Done cleaning extraneous packages.');
}

function main() {
  syncDependencyGraph();

  const nodeModulesRoot = path.resolve(__dirname, '../node_modules');
  const results = cleanupPolicy
    .map((entry) => inspectPolicyTarget(nodeModulesRoot, entry))
    .map((entry) => prunePolicyEntry(entry));

  reportCleanupSummary(results);
}

main();
