#!/bin/bash

set -e

APP_NAME="OpenCode Office Add-in"
APP_DIR="/Applications/${APP_NAME}.app"
LAUNCHAGENT_ID="com.opencode.office-addin"
MANIFEST_PATH="${APP_DIR}/Contents/Resources/manifest.xml"

user_data_dir() {
    printf '%s/Library/Application Support/OpenCode Office Add-in\n' "$1"
}

remove_generated_certificate() {
    local user_home="$1"
    local cert_dir
    local thumbprint_file

    cert_dir="$(user_data_dir "$user_home")/certs"
    thumbprint_file="$cert_dir/thumbprint.txt"

    if [ -f "$thumbprint_file" ]; then
        local sha1_hex
        sha1_hex="$(tr -d '\n\r' < "$thumbprint_file" | tr '[:lower:]' '[:upper:]')"
        if [ -n "$sha1_hex" ]; then
            security delete-certificate -Z "$sha1_hex" /Library/Keychains/System.keychain >/dev/null 2>&1 || true
        fi
    fi

    rm -rf "$(user_data_dir "$user_home")"
}

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

run_uninstall_action() {
    local install_user="$1"
    local user_home="$2"
    local action_name="$3"

    case "$action_name" in
        stop-launch-agent)
            local launch_agent_path="${user_home}/Library/LaunchAgents/${LAUNCHAGENT_ID}.plist"
            if [ -f "$launch_agent_path" ]; then
                sudo -u "$install_user" launchctl unload "$launch_agent_path" 2>/dev/null || true
                rm -f "$launch_agent_path"
            fi
            ;;
        kill-app)
            pkill -f "$APP_NAME" 2>/dev/null || true
            pkill -f "opencode-office-server" 2>/dev/null || true
            pkill -f "copilot-office-server" 2>/dev/null || true
            ;;
        remove-app)
            rm -rf "$APP_DIR"
            ;;
    esac
}

main() {
    local install_user
    local user_home
    local target_manifest_id

    echo "Uninstalling OpenCode Office Add-in..."
    install_user="$(resolve_install_user)"
    user_home="$(resolve_user_home "$install_user")"
    target_manifest_id="$(manifest_id "$MANIFEST_PATH")"

    remove_manifest_registrations "$user_home" "$target_manifest_id"
    run_uninstall_action "$install_user" "$user_home" stop-launch-agent
    run_uninstall_action "$install_user" "$user_home" kill-app

    echo "Removing application..."
    run_uninstall_action "$install_user" "$user_home" remove-app
    remove_generated_certificate "$user_home"

    echo ""
    echo "✓ OpenCode Office Add-in has been uninstalled."
    echo "Generated localhost certificate material and trust entries were removed."
}

main "$@"
