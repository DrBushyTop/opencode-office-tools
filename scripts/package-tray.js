#!/usr/bin/env node
const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const os = require('os');

function getPackagingContext() {
  const rootDir = path.resolve(__dirname, '..');
  const version = require(path.join(rootDir, 'package.json')).version;
  const platform = os.platform();
  const stageDir = path.join(rootDir, 'build', 'tray-package');
  const electronOutputDir = path.join(rootDir, 'build', 'electron');
  const zipName = platform === 'darwin'
    ? `opencode-office-addin-macos-v${version}.zip`
    : `opencode-office-addin-windows-v${version}.zip`;

  return {
    platform,
    rootDir,
    version,
    stageDir,
    electronOutputDir,
    zipName,
    zipPath: path.join(rootDir, 'build', zipName),
  };
}

function resetStageDirectory(stageDir) {
  if (fs.existsSync(stageDir)) {
    fs.rmSync(stageDir, { recursive: true });
  }
  fs.mkdirSync(stageDir, { recursive: true });
}

function runDesktopBuild(context) {
  console.log(`Packaging tray app for ${context.platform}...`);
  console.log('Building Electron app...');
  execSync('bun run clean:extraneous && bun run build', { cwd: context.rootDir, stdio: 'inherit' });

  if (context.platform === 'darwin') {
    execSync('bunx electron-builder --mac --dir', { cwd: context.rootDir, stdio: 'inherit' });
    return;
  }

  if (context.platform === 'win32') {
    execSync('bunx electron-builder --win --dir', { cwd: context.rootDir, stdio: 'inherit' });
    return;
  }

  throw new Error(`Unsupported platform: ${context.platform}`);
}

function locateBuiltBundle(context) {
  if (context.platform === 'darwin') {
    const appName = 'OpenCode Office Add-in.app';
    const candidates = [
      path.join(context.electronOutputDir, 'mac-arm64', appName),
      path.join(context.electronOutputDir, 'mac', appName),
    ];
    const sourceApp = candidates.find((candidate) => fs.existsSync(candidate));
    if (!sourceApp) {
      throw new Error('Could not find built macOS app');
    }

    return { source: sourceApp, registerScript: 'register.sh' };
  }

  const unpackedDir = path.join(context.electronOutputDir, 'win-unpacked');
  if (!fs.existsSync(unpackedDir)) {
    throw new Error('Could not find built Windows app');
  }

  return { source: unpackedDir, registerScript: 'register.ps1' };
}

function stageReleaseNotes(context) {
  fs.copyFileSync(
    path.join(context.rootDir, 'installer', 'GETTING_STARTED_RELEASE.md'),
    path.join(context.stageDir, 'GETTING_STARTED.md'),
  );
  console.log('Copied GETTING_STARTED.md');
}

function stageBundle(context, bundle) {
  if (context.platform === 'darwin') {
    console.log(`Copying ${path.basename(bundle.source)}...`);
    execSync(`cp -R "${bundle.source}" "${context.stageDir}/"`, { stdio: 'inherit' });
  } else {
    console.log('Copying Windows app...');
    execSync(`xcopy "${bundle.source}\\*" "${context.stageDir}\\" /E /I /Y`, { stdio: 'inherit' });
  }

  const sourceScript = path.join(context.rootDir, bundle.registerScript);
  const targetScript = path.join(context.stageDir, bundle.registerScript);
  fs.copyFileSync(sourceScript, targetScript);
  if (bundle.registerScript.endsWith('.sh')) {
    fs.chmodSync(targetScript, 0o755);
  }
  console.log(`Copied ${bundle.registerScript}`);
}

function createReleaseArchive(context) {
  if (fs.existsSync(context.zipPath)) {
    fs.unlinkSync(context.zipPath);
  }

  console.log(`Creating ${context.zipName}...`);
  if (context.platform === 'darwin') {
    execSync(`ditto -c -k --sequesterRsrc "${context.stageDir}" "${context.zipPath}"`, { stdio: 'inherit' });
  } else {
    execSync(
      `powershell -Command "Compress-Archive -Path '${context.stageDir}\\*' -DestinationPath '${context.zipPath}' -Force"`,
      { stdio: 'inherit' },
    );
  }

  console.log(`\nPackage created: build/${context.zipName}`);
}

function main() {
  const context = getPackagingContext();
  resetStageDirectory(context.stageDir);
  runDesktopBuild(context);
  stageReleaseNotes(context);
  stageBundle(context, locateBuiltBundle(context));
  createReleaseArchive(context);
}

main();
