#!/usr/bin/env node
const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const os = require('os');

function createReleaseProfile() {
  const projectRoot = path.resolve(__dirname, '..');
  const version = require(path.join(projectRoot, 'package.json')).version;
  const platform = os.platform();
  const stageRoot = path.join(projectRoot, 'build', 'tray-package');
  const archiveName = platform === 'darwin'
    ? `opencode-office-addin-macos-v${version}.zip`
    : `opencode-office-addin-windows-v${version}.zip`;

  return {
    projectRoot,
    platform,
    version,
    electronOutputRoot: path.join(projectRoot, 'build', 'electron'),
    stageRoot,
    archiveName,
    archivePath: path.join(projectRoot, 'build', archiveName),
    guideSource: path.join(projectRoot, 'installer', 'GETTING_STARTED_RELEASE.md'),
  };
}

function resetReleaseWorkspace(stageRoot) {
  if (fs.existsSync(stageRoot)) {
    fs.rmSync(stageRoot, { recursive: true });
  }
  fs.mkdirSync(stageRoot, { recursive: true });
}

function resolvePlatformBundle(profile) {
  if (profile.platform === 'darwin') {
    const appName = 'OpenCode Office Add-in.app';
    const candidates = [
      path.join(profile.electronOutputRoot, 'mac-arm64', appName),
      path.join(profile.electronOutputRoot, 'mac', appName),
    ];
    const appBundle = candidates.find((candidate) => fs.existsSync(candidate));
    if (!appBundle) {
      throw new Error('Could not find built macOS app');
    }

    return {
      buildCommand: 'bunx electron-builder --mac --dir',
      bundlePath: appBundle,
      bundleLabel: path.basename(appBundle),
      registrationScript: 'register.sh',
      copyBundle(stageRoot) {
        execSync(`cp -R "${appBundle}" "${stageRoot}/"`, { stdio: 'inherit' });
      },
      archive(stageRoot, archivePath) {
        execSync(`ditto -c -k --sequesterRsrc "${stageRoot}" "${archivePath}"`, { stdio: 'inherit' });
      },
    };
  }

  if (profile.platform === 'win32') {
    const unpackedDir = path.join(profile.electronOutputRoot, 'win-unpacked');
    if (!fs.existsSync(unpackedDir)) {
      throw new Error('Could not find built Windows app');
    }

    return {
      buildCommand: 'bunx electron-builder --win --dir',
      bundlePath: unpackedDir,
      bundleLabel: 'Windows app',
      registrationScript: 'register.ps1',
      copyBundle(stageRoot) {
        execSync(`xcopy "${unpackedDir}\\*" "${stageRoot}\\" /E /I /Y`, { stdio: 'inherit' });
      },
      archive(stageRoot, archivePath) {
        execSync(`powershell -Command "Compress-Archive -Path '${stageRoot}\\*' -DestinationPath '${archivePath}' -Force"`, { stdio: 'inherit' });
      },
    };
  }

  throw new Error(`Unsupported platform: ${profile.platform}`);
}

function buildDesktopRuntime(profile, platformBundle) {
  console.log(`Packaging tray app for ${profile.platform}...`);
  console.log('Preparing desktop runtime...');
  execSync('bun run clean:extraneous && bun run build', { cwd: profile.projectRoot, stdio: 'inherit' });
  execSync(platformBundle.buildCommand, { cwd: profile.projectRoot, stdio: 'inherit' });
}

function stageReleaseAssets(profile, platformBundle) {
  fs.copyFileSync(profile.guideSource, path.join(profile.stageRoot, 'GETTING_STARTED.md'));
  console.log('Staged GETTING_STARTED.md');

  const sourceScript = path.join(profile.projectRoot, platformBundle.registrationScript);
  const targetScript = path.join(profile.stageRoot, platformBundle.registrationScript);
  fs.copyFileSync(sourceScript, targetScript);
  if (platformBundle.registrationScript.endsWith('.sh')) {
    fs.chmodSync(targetScript, 0o755);
  }
  console.log(`Staged ${platformBundle.registrationScript}`);

  console.log(`Staging ${platformBundle.bundleLabel}...`);
  platformBundle.copyBundle(profile.stageRoot);
}

function writeArchive(profile, platformBundle) {
  if (fs.existsSync(profile.archivePath)) {
    fs.unlinkSync(profile.archivePath);
  }

  console.log(`Creating ${profile.archiveName}...`);
  platformBundle.archive(profile.stageRoot, profile.archivePath);
  console.log(`\nPackage created: build/${profile.archiveName}`);
}

function main() {
  const profile = createReleaseProfile();
  const platformBundle = resolvePlatformBundle(profile);
  resetReleaseWorkspace(profile.stageRoot);
  buildDesktopRuntime(profile, platformBundle);
  stageReleaseAssets(profile, platformBundle);
  writeArchive(profile, platformBundle);
}

main();
