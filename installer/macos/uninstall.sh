#!/bin/bash
# Uninstall script for macOS
# Run: sudo ./uninstall.sh

APP_NAME="OpenCode Office Add-in"
APP_DIR="/Applications/$APP_NAME.app"
LAUNCHAGENT="com.opencode.office-addin"
MANIFEST_FILENAME="opencode-office-addin.xml"
MANIFEST_PATH="$APP_DIR/Contents/Resources/manifest.xml"

manifest_id() {
    local manifest_path="$1"

    if [ ! -f "$manifest_path" ]; then
        return 1
    fi

    if command -v xmllint >/dev/null 2>&1; then
        xmllint --xpath 'string(/*[local-name()="OfficeApp"]/*[local-name()="Id"][1])' "$manifest_path" 2>/dev/null
        return 0
    fi

    grep -o '<Id>[^<]*</Id>' "$manifest_path" | sed -E 's#</?Id>##g' | head -n 1
}

remove_matching_manifests() {
    local wef_dir="$1"
    local target_manifest_id="$2"

    [ -d "$wef_dir" ] || return 0

    for existing_manifest in "$wef_dir"/*.xml; do
        [ -e "$existing_manifest" ] || continue

        local existing_id
        existing_id="$(manifest_id "$existing_manifest")"
        if [ -n "$target_manifest_id" ] && [ "$existing_id" = "$target_manifest_id" ]; then
            rm -f "$existing_manifest"
            continue
        fi

        if [ -z "$target_manifest_id" ] && [ "$(basename "$existing_manifest")" = "$MANIFEST_FILENAME" ]; then
            rm -f "$existing_manifest"
        fi
    done
}

echo "Uninstalling OpenCode Office Add-in..."

# Get the current user
if [ -n "$SUDO_USER" ]; then
    INSTALL_USER="$SUDO_USER"
else
    INSTALL_USER=$(stat -f "%Su" /dev/console)
fi

USER_HOME=$(dscl . -read /Users/$INSTALL_USER NFSHomeDirectory | awk '{print $2}')
TARGET_MANIFEST_ID="$(manifest_id "$MANIFEST_PATH")"

# Stop the service
echo "Stopping service..."
LAUNCHAGENT_PATH="$USER_HOME/Library/LaunchAgents/$LAUNCHAGENT.plist"
if [ -f "$LAUNCHAGENT_PATH" ]; then
    sudo -u $INSTALL_USER launchctl unload "$LAUNCHAGENT_PATH" 2>/dev/null || true
    rm -f "$LAUNCHAGENT_PATH"
fi

# Kill any running Electron app
pkill -f "$APP_NAME" 2>/dev/null || true

# Also kill any old standalone server process (from previous versions)
pkill -f "opencode-office-server" 2>/dev/null || true

# Remove add-in registrations
echo "Removing add-in registrations..."
WORD_WEF="$USER_HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef"
PPT_WEF="$USER_HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
EXCEL_WEF="$USER_HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
ONENOTE_WEF="$USER_HOME/Library/Containers/com.microsoft.onenote.mac/Data/Documents/wef"

for WEF_DIR in "$WORD_WEF" "$PPT_WEF" "$EXCEL_WEF" "$ONENOTE_WEF"; do
    remove_matching_manifests "$WEF_DIR" "$TARGET_MANIFEST_ID"
done

# Remove application directory
echo "Removing application..."
rm -rf "$APP_DIR"

echo ""
echo "✓ OpenCode Office Add-in has been uninstalled."
echo ""
echo "Note: The SSL certificate remains in your keychain."
echo "To remove it manually:"
echo "  1. Open Keychain Access"
echo "  2. Search for 'localhost'"
echo "  3. Delete the certificate"
