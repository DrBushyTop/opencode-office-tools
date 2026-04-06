const sharp = require('sharp');
const path = require('path');
const fs = require('fs');
const toIco = require('to-ico');
const { execSync } = require('child_process');

const pngSizes = [16, 32, 64, 80, 128, 256, 512, 1024];
const rootDir = path.resolve(__dirname, '..');
const sourceLogo = path.join(rootDir, 'logo.svg');
const webIconDir = path.join(rootDir, 'src', 'ui', 'public');
const windowsInstallerDir = path.join(rootDir, 'installer', 'windows');
const macInstallerDir = path.join(rootDir, 'installer', 'macos');
const assetsDir = path.join(rootDir, 'assets');

function ensureDirectory(targetDir) {
  if (!fs.existsSync(targetDir)) {
    fs.mkdirSync(targetDir, { recursive: true });
  }
}

async function renderPngIcon(size, targetFile) {
  await sharp(sourceLogo)
    .resize(size, size, {
      fit: 'contain',
      background: { r: 0, g: 0, b: 0, alpha: 0 },
    })
    .png()
    .toFile(targetFile);
}

async function buildWebIcons() {
  ensureDirectory(webIconDir);
  console.log('Generating icons from logo.svg...');

  const outputFiles = [];
  for (const size of pngSizes) {
    const targetFile = path.join(webIconDir, `icon-${size}.png`);
    await renderPngIcon(size, targetFile);
    outputFiles.push(targetFile);
    console.log(`  ✓ icon-${size}.png`);
  }

  console.log(`\nPNG icons saved to: ${webIconDir}`);
  return outputFiles;
}

async function buildWindowsIco() {
  console.log('\nGenerating Windows .ico file...');
  try {
    const icoSizes = [16, 32, 64, 128, 256];
    const pngBuffers = icoSizes.map((size) => fs.readFileSync(path.join(webIconDir, `icon-${size}.png`)));
    const icoBuffer = await toIco(pngBuffers);
    const installerIcoPath = path.join(windowsInstallerDir, 'app.ico');
    const trayIcoPath = path.join(assetsDir, 'tray-icon.ico');
    fs.writeFileSync(installerIcoPath, icoBuffer);
    fs.writeFileSync(trayIcoPath, icoBuffer);
    console.log(`  ✓ app.ico saved to: ${installerIcoPath}`);
    console.log(`  ✓ tray-icon.ico saved to: ${trayIcoPath}`);
  } catch (error) {
    console.log(`  ⚠ Could not generate .ico: ${error.message}`);
  }
}

function copyMacInstallerPng() {
  const destination = path.join(macInstallerDir, 'icon.png');
  fs.copyFileSync(path.join(webIconDir, 'icon-256.png'), destination);
  console.log(`  ✓ icon.png copied to: ${destination}`);
}

async function buildMacIcns() {
  console.log('\nGenerating macOS .icns file...');
  if (process.platform !== 'darwin') {
    console.log('  ⚠ Skipping .icns generation (only works on macOS)');
    return;
  }

  const iconsetDir = path.join(macInstallerDir, 'icon.iconset');
  const iconsetEntries = [
    ['icon_16x16.png', 16],
    ['icon_16x16@2x.png', 32],
    ['icon_32x32.png', 32],
    ['icon_32x32@2x.png', 64],
    ['icon_128x128.png', 128],
    ['icon_128x128@2x.png', 256],
    ['icon_256x256.png', 256],
    ['icon_256x256@2x.png', 512],
    ['icon_512x512.png', 512],
    ['icon_512x512@2x.png', 1024],
  ];

  try {
    ensureDirectory(iconsetDir);
    for (const [filename, size] of iconsetEntries) {
      fs.copyFileSync(path.join(webIconDir, `icon-${size}.png`), path.join(iconsetDir, filename));
    }

    const icnsPath = path.join(macInstallerDir, 'icon.icns');
    execSync(`iconutil -c icns "${iconsetDir}" -o "${icnsPath}"`);
    console.log(`  ✓ icon.icns saved to: ${icnsPath}`);
    fs.rmSync(iconsetDir, { recursive: true });
  } catch (error) {
    console.log(`  ⚠ Could not generate .icns: ${error.message}`);
    console.log('    (This only works on macOS with iconutil installed)');
  }
}

async function buildTrayIcons() {
  console.log('\nGenerating tray icons...');
  try {
    const trayTargets = [
      { file: 'tray-icon.png', size: 16, template: false },
      { file: 'tray-icon@2x.png', size: 32, template: false },
      { file: 'tray-iconTemplate.png', size: 16, template: true },
      { file: 'tray-iconTemplate@2x.png', size: 32, template: true },
    ];

    for (const target of trayTargets) {
      let pipeline = sharp(sourceLogo).resize(target.size, target.size, {
        fit: 'contain',
        background: { r: 0, g: 0, b: 0, alpha: 0 },
      });
      if (target.template) {
        pipeline = pipeline.grayscale().threshold();
      }
      await pipeline.png().toFile(path.join(assetsDir, target.file));
      console.log(`  ✓ ${target.file}`);
    }
  } catch (error) {
    console.log(`  ⚠ Could not generate tray icons: ${error.message}`);
  }
}

async function main() {
  await buildWebIcons();
  await buildWindowsIco();
  copyMacInstallerPng();
  await buildMacIcns();
  await buildTrayIcons();
}

main().catch((error) => {
  console.error('Error generating icons:', error);
  process.exit(1);
});
