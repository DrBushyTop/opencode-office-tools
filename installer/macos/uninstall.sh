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
        return 0
    fi

    if command -v xmllint >/dev/null 2>&1; then
        xmllint --xpath 'string(/*[local-name()="OfficeApp"]/*[local-name()="Id"][1])' "$manifest_path" 2>/dev/null || true
        return 0
    fi

    grep -o '<Id>[^<]*</Id>' "$manifest_path" 2>/dev/null | sed -E 's#</?Id>##g' | head -n 1 || true
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

    # osascript "with administrator privileges" runs as root without SUDO_USER;
    # stat the console device to find the GUI-session owner.
    local console_user
    console_user="$(stat -f "%Su" /dev/console 2>/dev/null || true)"
    if [ -n "$console_user" ] && [ "$console_user" != "root" ]; then
        printf '%s\n' "$console_user"
        return
    fi

    # Last resort: the user who owns /Applications/<app>.app
    if [ -d "$APP_DIR" ]; then
        console_user="$(stat -f "%Su" "$APP_DIR" 2>/dev/null || true)"
        if [ -n "$console_user" ] && [ "$console_user" != "root" ]; then
            printf '%s\n' "$console_user"
            return
        fi
    fi

    echo "root"
}

resolve_user_home() {
    local home_dir
    home_dir="$(dscl . -read "/Users/$1" NFSHomeDirectory 2>/dev/null | awk '{print $2}')"
    if [ -z "$home_dir" ]; then
        # Fallback: use eval to expand ~user
        home_dir="$(eval echo "~$1" 2>/dev/null || echo "/Users/$1")"
    fi
    printf '%s\n' "$home_dir"
}

remove_manifest_registrations() {
    local user_home="$1"
    local target_manifest_id="$2"
    local host_dirs=(
        "$user_home/Library/Containers/com.microsoft.Word/Data/Documents/wef"
        "$user_home/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
        "$user_home/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
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
                local uid
                uid="$(id -u "$install_user" 2>/dev/null || echo "")"
                if [ -n "$uid" ] && [ "$uid" != "0" ]; then
                    # Prefer bootout (modern launchctl); fall back to unload
                    launchctl bootout "gui/${uid}/${LAUNCHAGENT_ID}" 2>/dev/null || \
                        sudo -u "$install_user" launchctl unload "$launch_agent_path" 2>/dev/null || true
                fi
                rm -f "$launch_agent_path"
            fi
            ;;
        kill-app)
            # Match the Electron binary path (Contents/MacOS), NOT the broad
            # app name — pkill -f "$APP_NAME" would also match this script's
            # own process tree (bash, osascript) and kill the uninstaller.
            pkill -f "${APP_DIR}/Contents/MacOS" 2>/dev/null || true
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

    echo "Removing package receipt..."
    pkgutil --forget com.opencode.office-addin >/dev/null 2>&1 || true

    echo ""
    echo "✓ OpenCode Office Add-in has been uninstalled."
    echo "Generated localhost certificate material and trust entries were removed."
}

main "$@"
