#!/bin/bash

set -euo pipefail

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
ROOT_DIR="${SCRIPT_DIR}/../.."
BUILD_DIR="${ROOT_DIR}/build/macos"
ELECTRON_OUTPUT_DIR="${ROOT_DIR}/build/electron"
APP_NAME="OpenCode Office Add-in"
APP_IDENTIFIER="com.opencode.office-addin"
APP_VERSION="1.0.0"

ensure_icon_assets() {
  if [ -f "${SCRIPT_DIR}/icon.icns" ]; then
    return
  fi

  echo "Generating icons..."
  (
    cd "${ROOT_DIR}"
    bun run build:icons
  )
}

build_desktop_bundle() {
  echo "Building Electron app..."
  (
    cd "${ROOT_DIR}"
    bun run clean:extraneous
    bun run build
    bunx electron-builder --mac
  )
}

locate_app_bundle() {
  local candidate
  for candidate in \
    "${ELECTRON_OUTPUT_DIR}/mac-arm64/${APP_NAME}.app" \
    "${ELECTRON_OUTPUT_DIR}/mac/${APP_NAME}.app"; do
    if [ -d "$candidate" ]; then
      printf '%s\n' "$candidate"
      return 0
    fi
  done

  echo "Error: Could not find built app in ${ELECTRON_OUTPUT_DIR}" >&2
  exit 1
}

prepare_stage_layout() {
  rm -rf "${BUILD_DIR}"
  mkdir -p "${BUILD_DIR}/component-root/Applications"
  mkdir -p "${BUILD_DIR}/script-root"
  mkdir -p "${BUILD_DIR}/resource-root"
}

copy_payload_assets() {
  local app_bundle="$1"

  cp -R "$app_bundle" "${BUILD_DIR}/component-root/Applications/"
  cp "${SCRIPT_DIR}/launchagent/com.opencode.office-addin.plist" "${BUILD_DIR}/component-root/Applications/${APP_NAME}.app/Contents/Resources/"
  cp "${SCRIPT_DIR}/scripts/preinstall" "${BUILD_DIR}/script-root/"
  cp "${SCRIPT_DIR}/scripts/postinstall" "${BUILD_DIR}/script-root/"
  chmod +x "${BUILD_DIR}/script-root/preinstall"
  chmod +x "${BUILD_DIR}/script-root/postinstall"
}

write_installer_documents() {
  cat > "${BUILD_DIR}/resource-root/welcome.html" <<'EOF'
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, sans-serif; padding: 20px; }
        h1 { color: #24292f; }
        p { color: #57606a; line-height: 1.5; }
        ul { color: #57606a; }
    </style>
</head>
<body>
    <h1>OpenCode Office Add-in</h1>
    <p>This installer will set up the OpenCode Office Add-in on your Mac.</p>
    <p>The installer will:</p>
    <ul>
        <li>Install the add-in application to your Applications folder</li>
        <li>Register the add-in with Word, PowerPoint, Excel, and OneNote</li>
        <li>Configure the service to start automatically at login</li>
        <li>Add a menu bar icon for easy access</li>
    </ul>
    <p>Click Continue to proceed with the installation.</p>
</body>
</html>
EOF

  cat > "${BUILD_DIR}/resource-root/conclusion.html" <<'EOF'
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, sans-serif; padding: 20px; }
        h1 { color: #24292f; }
        p { color: #57606a; line-height: 1.5; }
        .success { color: #1a7f37; font-weight: 600; }
    </style>
</head>
<body>
    <h1>Installation Complete!</h1>
    <p class="success">✓ OpenCode Office Add-in has been installed successfully.</p>
    <p>The add-in is now running in your menu bar.</p>
    <p><strong>Next steps:</strong></p>
    <ol>
        <li>Look for the OpenCode icon in your menu bar</li>
        <li>Open Word, PowerPoint, Excel, or OneNote</li>
        <li>Find the "OpenCode" button on the Home ribbon</li>
        <li>Click the button to open the OpenCode panel</li>
    </ol>
    <p>The app will start automatically when you log in.</p>
</body>
</html>
EOF
}

write_distribution_spec() {
  cat > "${BUILD_DIR}/distribution.xml" <<EOF
<?xml version="1.0" encoding="utf-8"?>
<installer-gui-script minSpecVersion="2">
    <title>${APP_NAME}</title>
    <organization>${APP_IDENTIFIER}</organization>
    <domains enable_localSystem="true" enable_currentUserHome="false"/>
    <options customize="never" require-scripts="true" rootVolumeOnly="true"/>
    <welcome file="welcome.html"/>
    <conclusion file="conclusion.html"/>
    <pkg-ref id="${APP_IDENTIFIER}"/>
    <choices-outline>
        <line choice="default">
            <line choice="${APP_IDENTIFIER}"/>
        </line>
    </choices-outline>
    <choice id="default"/>
    <choice id="${APP_IDENTIFIER}" visible="false">
        <pkg-ref id="${APP_IDENTIFIER}"/>
    </choice>
    <pkg-ref id="${APP_IDENTIFIER}" version="${APP_VERSION}" onConclusion="none">OpenCodeOfficeAddin-component.pkg</pkg-ref>
</installer-gui-script>
EOF
}

build_component_pkg() {
  pkgbuild \
    --root "${BUILD_DIR}/component-root" \
    --scripts "${BUILD_DIR}/script-root" \
    --identifier "${APP_IDENTIFIER}" \
    --version "${APP_VERSION}" \
    --install-location "/" \
    "${BUILD_DIR}/OpenCodeOfficeAddin-component.pkg"
}

build_distribution_pkg() {
  productbuild \
    --distribution "${BUILD_DIR}/distribution.xml" \
    --resources "${BUILD_DIR}/resource-root" \
    --package-path "${BUILD_DIR}" \
    "${BUILD_DIR}/OpenCodeOfficeAddin-${APP_VERSION}.pkg"
}

cleanup_stage_layout() {
  rm -f "${BUILD_DIR}/OpenCodeOfficeAddin-component.pkg"
  rm -f "${BUILD_DIR}/distribution.xml"
  rm -rf "${BUILD_DIR}/component-root"
  rm -rf "${BUILD_DIR}/script-root"
  rm -rf "${BUILD_DIR}/resource-root"
}

main() {
  echo "Building macOS installer..."
  ensure_icon_assets
  build_desktop_bundle

  local app_bundle
  app_bundle="$(locate_app_bundle)"
  echo "App location: ${app_bundle}"

  prepare_stage_layout
  copy_payload_assets "$app_bundle"
  write_installer_documents
  write_distribution_spec

  echo "Building component package..."
  build_component_pkg
  echo "Building distribution package..."
  build_distribution_pkg
  cleanup_stage_layout

  echo ""
  echo "✓ macOS installer built successfully!"
  echo "  Output: ${BUILD_DIR}/OpenCodeOfficeAddin-${APP_VERSION}.pkg"
}

main "$@"
