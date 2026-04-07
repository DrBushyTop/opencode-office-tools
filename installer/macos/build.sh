#!/bin/bash

set -euo pipefail

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
ROOT_DIR="${SCRIPT_DIR}/../.."
BUILD_DIR="${ROOT_DIR}/build/macos"
ELECTRON_OUTPUT_DIR="${ROOT_DIR}/build/electron"
APP_NAME="OpenCode Office Add-in"
APP_IDENTIFIER="com.opencode.office-addin"
APP_VERSION="1.0.0"
TARGET_ARCH="${TARGET_ARCH:-}"

package_filename() {
  if [ -n "${TARGET_ARCH}" ]; then
    printf 'OpenCodeOfficeAddin-%s-%s.pkg\n' "${APP_VERSION}" "${TARGET_ARCH}"
    return
  fi

  printf 'OpenCodeOfficeAddin-%s.pkg\n' "${APP_VERSION}"
}

ensure_icon_assets() {
  if [ -f "${SCRIPT_DIR}/icon.icns" ]; then
    return
  fi

  echo "Generating installer icons..."
  (
    cd "${ROOT_DIR}"
    bun run build:icons
  )
}

build_desktop_bundle() {
  echo "Building desktop bundle..."
  (
    cd "${ROOT_DIR}"
    bun run clean:extraneous
    bun run build
    if [ -n "${TARGET_ARCH}" ]; then
      bunx electron-builder --dir --mac --${TARGET_ARCH}
    else
      bunx electron-builder --dir --mac
    fi
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
  cp "${SCRIPT_DIR}/uninstall.sh" "${BUILD_DIR}/component-root/Applications/${APP_NAME}.app/Contents/Resources/"
  cp "${SCRIPT_DIR}/scripts/preinstall" "${BUILD_DIR}/script-root/"
  cp "${SCRIPT_DIR}/scripts/postinstall" "${BUILD_DIR}/script-root/"
  chmod +x "${BUILD_DIR}/component-root/Applications/${APP_NAME}.app/Contents/Resources/uninstall.sh"
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
    <p>This installer prepares OpenCode Office Add-in for local Office sideloading on your Mac.</p>
    <p>During setup, it will:</p>
    <ul>
        <li>Copy the desktop app into your Applications folder</li>
        <li>Refresh the OpenCode sideload manifest for Word, PowerPoint, Excel, and OneNote</li>
        <li>Configure the helper app to start automatically at login</li>
        <li>Expose a menu bar entry for status and troubleshooting</li>
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
    <h1>Installation Complete</h1>
    <p class="success">✓ OpenCode Office Add-in is installed and the local helper has been started.</p>
    <p>You should now see the app in your menu bar.</p>
    <p><strong>Next steps:</strong></p>
    <ol>
        <li>Confirm the OpenCode icon is visible in your menu bar</li>
        <li>Open Word, PowerPoint, Excel, or OneNote</li>
        <li>Open the "OpenCode" command from the Home ribbon</li>
        <li>Wait for the task pane to connect to the local helper</li>
    </ol>
    <p>The app will relaunch automatically the next time you sign in.</p>
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
  local output_name
  output_name="$(package_filename)"

  productbuild \
    --distribution "${BUILD_DIR}/distribution.xml" \
    --resources "${BUILD_DIR}/resource-root" \
    --package-path "${BUILD_DIR}" \
    "${BUILD_DIR}/${output_name}"
}

cleanup_stage_layout() {
  rm -f "${BUILD_DIR}/OpenCodeOfficeAddin-component.pkg"
  rm -f "${BUILD_DIR}/distribution.xml"
  rm -rf "${BUILD_DIR}/component-root"
  rm -rf "${BUILD_DIR}/script-root"
  rm -rf "${BUILD_DIR}/resource-root"
}

main() {
  echo "Building macOS installer package..."
  ensure_icon_assets
  build_desktop_bundle

  local app_bundle
  app_bundle="$(locate_app_bundle)"
  echo "Resolved app bundle: ${app_bundle}"

  prepare_stage_layout
  copy_payload_assets "$app_bundle"
  write_installer_documents
  write_distribution_spec

  echo "Creating component package..."
  build_component_pkg
  echo "Creating distribution package..."
  build_distribution_pkg
  cleanup_stage_layout

  echo ""
  echo "✓ macOS installer package built successfully"
  echo "  Output: ${BUILD_DIR}/$(package_filename)"
}

main "$@"
