#!/bin/bash

set -e

APP_NAME="OpenCode Office Add-in"
APP_DIR="/Applications/${APP_NAME}.app"
LAUNCHAGENT_ID="com.opencode.office-addin"
MANIFEST_PATH="${APP_DIR}/Contents/Resources/manifest.xml"

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

is_opencode_manifest_name() {
    local filename="$1"
    [ "$filename" = "manifest.xml" ] || [ "$filename" = "opencode-office-addin.xml" ]
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

        if [ -z "$target_manifest_id" ] && is_opencode_manifest_name "$(basename "$existing_manifest")"; then
            rm -f "$existing_manifest"
        fi
    done
}

resolve_install_user() {
    if [ -n "${SUDO_USER:-}" ]; then
        printf '%s\n' "$SUDO_USER"
        return
    fi

    stat -f "%Su" /dev/console
}

resolve_user_home() {
    dscl . -read "/Users/$1" NFSHomeDirectory | awk '{print $2}'
}

stop_background_items() {
    local install_user="$1"
    local user_home="$2"
    local launch_agent_path="${user_home}/Library/LaunchAgents/${LAUNCHAGENT_ID}.plist"

    echo "Stopping service..."
    if [ -f "$launch_agent_path" ]; then
        sudo -u "$install_user" launchctl unload "$launch_agent_path" 2>/dev/null || true
        rm -f "$launch_agent_path"
    fi

    pkill -f "$APP_NAME" 2>/dev/null || true
    pkill -f "opencode-office-server" 2>/dev/null || true
    pkill -f "copilot-office-server" 2>/dev/null || true
}

remove_manifest_registrations() {
    local user_home="$1"
    local target_manifest_id="$2"
    local host_dirs=(
        "$user_home/Library/Containers/com.microsoft.Word/Data/Documents/wef"
        "$user_home/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
        "$user_home/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
        "$user_home/Library/Containers/com.microsoft.onenote.mac/Data/Documents/wef"
    )

    echo "Removing add-in registrations..."
    local wef_dir
    for wef_dir in "${host_dirs[@]}"; do
        remove_matching_manifests "$wef_dir" "$target_manifest_id"
    done
}

main() {
    local install_user
    local user_home
    local target_manifest_id

    echo "Uninstalling OpenCode Office Add-in..."
    install_user="$(resolve_install_user)"
    user_home="$(resolve_user_home "$install_user")"
    target_manifest_id="$(manifest_id "$MANIFEST_PATH")"

    stop_background_items "$install_user" "$user_home"
    remove_manifest_registrations "$user_home" "$target_manifest_id"

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
}

main "$@"
