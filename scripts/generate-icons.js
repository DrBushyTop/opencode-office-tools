const sharp = require('sharp');
const path = require('path');
const fs = require('fs');
const toIco = require('to-ico');
const { execSync } = require('child_process');

const rootDir = path.resolve(__dirname, '..');
const sourceLogo = path.join(rootDir, 'logo.png');
const webIconDir = path.join(rootDir, 'src', 'ui', 'public');
const windowsInstallerDir = path.join(rootDir, 'installer', 'windows');
const macInstallerDir = path.join(rootDir, 'installer', 'macos');
const assetsDir = path.join(rootDir, 'assets');

const pngSizes = [16, 32, 64, 80, 128, 256, 512, 1024];
const icnsLayout = [
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
const trayAssetPlan = [
  { file: 'tray-icon.png', size: 16, template: false },
  { file: 'tray-icon@2x.png', size: 32, template: false },
  { file: 'tray-iconTemplate.png', size: 16, template: true },
  { file: 'tray-iconTemplate@2x.png', size: 32, template: true },
];

function ensureDirectory(targetDir) {
  if (!fs.existsSync(targetDir)) {
    fs.mkdirSync(targetDir, { recursive: true });
  }
}

async function renderSourceCatalog() {
  // Trim transparent padding first, then resize to square at each target size.
  const trimmed = await sharp(sourceLogo).trim().toBuffer();
  const rendered = new Map();
  for (const size of pngSizes) {
    const buffer = await sharp(trimmed)
      .resize(size, size, {
        fit: 'contain',
        background: { r: 0, g: 0, b: 0, alpha: 0 },
      })
      .png()
      .toBuffer();
    rendered.set(size, buffer);
  }
  return rendered;
}

function bufferForSize(rendered, size) {
  const buffer = rendered.get(size);
  if (!buffer) {
    throw new Error(`Missing rendered icon buffer for size ${size}.`);
  }
  return buffer;
}

async function writeBrowserIcons(rendered) {
  ensureDirectory(webIconDir);
  console.log('Generating icons from logo.png...');

  for (const size of pngSizes) {
    const targetFile = path.join(webIconDir, `icon-${size}.png`);
    fs.writeFileSync(targetFile, bufferForSize(rendered, size));
    console.log(`  ✓ icon-${size}.png`);
  }

  console.log(`\nPNG icons saved to: ${webIconDir}`);
}

async function writeWindowsAssets(rendered) {
  console.log('\nGenerating Windows installer assets...');
  try {
    const icoBuffer = await toIco([16, 32, 64, 128, 256].map((size) => bufferForSize(rendered, size)));
    const installerIcoPath = path.join(windowsInstallerDir, 'app.ico');
    const trayIcoPath = path.join(assetsDir, 'tray-icon.ico');
    fs.writeFileSync(installerIcoPath, icoBuffer);
    fs.writeFileSync(trayIcoPath, icoBuffer);
    console.log(`  ✓ app.ico saved to: ${installerIcoPath}`);
    console.log(`  ✓ tray-icon.ico saved to: ${trayIcoPath}`);
  } catch (error) {
    console.log(`  ⚠ Could not generate Windows .ico asset: ${error.message}`);
  }
}

function writeMacInstallerPng(rendered) {
  const destination = path.join(macInstallerDir, 'icon.png');
  fs.writeFileSync(destination, bufferForSize(rendered, 256));
  console.log(`  ✓ icon.png copied to: ${destination}`);
}

async function writeTrayAssets(trimmed) {
  console.log('\nGenerating tray assets...');
  try {
    for (const asset of trayAssetPlan) {
      let pipeline = sharp(trimmed).resize(asset.size, asset.size, {
        fit: 'contain',
        background: { r: 0, g: 0, b: 0, alpha: 0 },
      });
      if (asset.template) {
        pipeline = pipeline.grayscale().threshold();
      }

      await pipeline.png().toFile(path.join(assetsDir, asset.file));
      console.log(`  ✓ ${asset.file}`);
    }
  } catch (error) {
    console.log(`  ⚠ Could not generate tray assets: ${error.message}`);
  }
}

async function writeMacIcns(rendered) {
  console.log('\nGenerating macOS installer assets...');
  if (process.platform !== 'darwin') {
    console.log('  ⚠ Skipping .icns generation (only works on macOS)');
    return;
  }

  const iconsetDir = path.join(macInstallerDir, 'icon.iconset');
  try {
    ensureDirectory(iconsetDir);
    for (const [filename, size] of icnsLayout) {
      fs.writeFileSync(path.join(iconsetDir, filename), bufferForSize(rendered, size));
    }

    const icnsPath = path.join(macInstallerDir, 'icon.icns');
    execSync(`iconutil -c icns "${iconsetDir}" -o "${icnsPath}"`);
    console.log(`  ✓ icon.icns saved to: ${icnsPath}`);
    fs.rmSync(iconsetDir, { recursive: true, force: true });
  } catch (error) {
    console.log(`  ⚠ Could not generate .icns: ${error.message}`);
    console.log('    (This only works on macOS with iconutil installed)');
  }
}

async function main() {
  // Trim transparent padding from source once, reuse for all outputs.
  const trimmed = await sharp(sourceLogo).trim().toBuffer();
  const renderedCatalog = await renderSourceCatalog();
  await writeBrowserIcons(renderedCatalog);
  await writeWindowsAssets(renderedCatalog);
  writeMacInstallerPng(renderedCatalog);
  await writeTrayAssets(trimmed);
  await writeMacIcns(renderedCatalog);
}

main().catch((error) => {
  console.error('Error generating icons:', error);
  process.exit(1);
});
