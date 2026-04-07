# Building Installers

This directory contains the packaging assets for the Windows and macOS desktop installers.

## Prerequisites

1. **OpenCode CLI** installed on the system (see [GETTING_STARTED.md](../GETTING_STARTED.md#installing-opencode))
2. **Node.js 20+** installed
3. **Bun dependencies** installed: `bun install`

> **Note:** The installers do **not** install the OpenCode CLI. It still needs to be installed separately before the add-in can serve requests.

### macOS-specific
- Xcode Command Line Tools: `xcode-select --install`

## Building

### Build for Current Platform

```bash
bun run build:installer
```

### macOS Only

```bash
bun run build:installer:mac
```

**Output:** `build/macos/*.pkg`

### Windows Only

```bash
bun run build:installer:win
```

**Output:** `build/electron/*.exe`

## What the Installers Do

### Windows Installer (.exe)
1. Installs the desktop app to `C:\Program Files\OpenCode Office Add-in\`
2. Bundles the built frontend and manifest
3. Generates a per-user localhost certificate under `%APPDATA%\OpenCode Office Add-in\certs`
4. Trusts that generated certificate in the current user's Root store
5. Refreshes the Office sideload manifest registration in the developer registry key
6. Creates a startup entry so the helper app can relaunch on sign-in
7. Starts the tray app immediately after install

### macOS Installer (.pkg)
1. Installs to `/Applications/OpenCode Office Add-in.app/`
2. Bundles the built frontend and manifest
3. Generates a per-user localhost certificate under `~/Library/Application Support/OpenCode Office Add-in/certs`
4. Trusts that generated certificate in the System keychain
5. Refreshes the sideload manifest in the Word, PowerPoint, and Excel WEF folders
6. Installs a LaunchAgent so the helper app can relaunch on sign-in
7. Starts the tray app immediately after install

## Uninstalling

### Windows
Use "Add or Remove Programs" in Windows Settings. The NSIS uninstaller removes the generated cert, app data, and Office registration.

### macOS
```bash
sudo /Applications/OpenCode\ Office\ Add-in.app/Contents/Resources/uninstall.sh
```

The uninstall script removes the generated certificate, LaunchAgent, packaged app, and Office sideload registrations.

## Code Signing (Optional)

For distribution outside your organization, you should sign the installers.

### Windows
Set the following environment variables before building:
- `CSC_LINK` - Path to your .pfx certificate file
- `CSC_KEY_PASSWORD` - Certificate password

Or use `signtool.exe` after building:
```powershell
signtool sign /f certificate.pfx /p password /t http://timestamp.digicert.com "build\electron\OpenCode Office Add-in Setup.exe"
```

### macOS
Set the following environment variables before building:
- `CSC_NAME` - Your Developer ID certificate name (e.g., "Developer ID Application: Your Name (TEAMID)")

Or sign the PKG after building:
```bash
productsign --sign "Developer ID Installer: Your Name (TEAMID)" "build/macos/OpenCodeOfficeAddin-1.0.0.pkg" "build/macos/OpenCodeOfficeAddin-1.0.0-signed.pkg"
```

## Troubleshooting

### Service not starting
- **Windows**: Check Task Scheduler or startup entries for "OpenCodeOfficeAddin"
- **macOS**: Check `launchctl list | grep opencode` and logs in `/tmp/opencode-office-addin.log`

### Add-in not appearing in Office
1. Ensure the service is running: visit https://localhost:52390 in browser
2. Restart the Office application
3. Check the manifest is registered:
   - **Windows**: `reg query "HKCU\Software\Microsoft\Office\16.0\WEF\Developer"`
   - **macOS**: Check `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/`

### Local certificate issues
1. Visit https://localhost:52390 in your browser
2. If you see a certificate warning, the cert isn't trusted
3. Re-run the installer so it can regenerate and trust a fresh localhost certificate
