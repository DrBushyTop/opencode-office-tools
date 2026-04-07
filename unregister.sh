#!/bin/bash

# Resolve script-relative resources once so repo and packaged layouts behave the same.
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
APP_PATH="$SCRIPT_DIR/OpenCode Office Add-in.app"
MANIFEST_FILENAME="manifest.xml"

resolve_user_home() {
    if [ -n "${HOME:-}" ]; then
        printf '%s\n' "$HOME"
        return
    fi

    dscl . -read "/Users/$(stat -f "%Su" /dev/console)" NFSHomeDirectory | awk '{print $2}'
}

packaged_user_data_dir() {
    printf '%s/OpenCode Office Add-in\n' "$(resolve_user_home)/Library/Application Support"
}

remove_packaged_certificate() {
    local cert_dir="$1"
    local thumbprint_file="$cert_dir/thumbprint.txt"

    if [ -f "$thumbprint_file" ]; then
        local sha1_hex
        sha1_hex="$(tr -d '\n\r' < "$thumbprint_file" | tr '[:lower:]' '[:upper:]')"
        if [ -n "$sha1_hex" ]; then
            security delete-certificate -Z "$sha1_hex" /Library/Keychains/System.keychain >/dev/null 2>&1 || true
        fi
    fi

    rm -rf "$cert_dir"
}

if [ -d "$APP_PATH" ]; then
    MANIFEST_PATH="$APP_PATH/Contents/Resources/manifest.xml"
else
    MANIFEST_PATH="$SCRIPT_DIR/manifest.xml"
fi

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
    local removed=1

    [ -d "$wef_dir" ] || return 1

    for existing_manifest in "$wef_dir"/*.xml; do
        [ -e "$existing_manifest" ] || continue

        local existing_id
        existing_id="$(manifest_id "$existing_manifest")"
        if [ -n "$target_manifest_id" ] && [ "$existing_id" = "$target_manifest_id" ]; then
            rm -f "$existing_manifest"
            removed=0
            continue
        fi

        if [ -z "$target_manifest_id" ] && is_opencode_manifest_name "$(basename "$existing_manifest")"; then
            rm -f "$existing_manifest"
            removed=0
        fi
    done

    return $removed
}

TARGET_MANIFEST_ID="$(manifest_id "$MANIFEST_PATH")"

echo -e "\033[36mRemoving OpenCode Office Add-in registration from macOS...\033[0m"
echo ""

# Office keeps sideload manifests in a per-host WEF directory.
WORD_WEF_DIR="$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef"
POWERPOINT_WEF_DIR="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
EXCEL_WEF_DIR="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"

# Remove any matching OpenCode registration from each host folder.
if [ -f "$WORD_WEF_DIR/$MANIFEST_FILENAME" ]; then
    rm "$WORD_WEF_DIR/$MANIFEST_FILENAME"
    echo -e "  \033[32m✓ Removed Word sideload registration\033[0m"
else
    if remove_matching_manifests "$WORD_WEF_DIR" "$TARGET_MANIFEST_ID"; then
        echo -e "  \033[32m✓ Removed Word sideload registration\033[0m"
    else
        echo -e "  \033[90m• No OpenCode registration found for Word\033[0m"
    fi
fi

if [ -f "$POWERPOINT_WEF_DIR/$MANIFEST_FILENAME" ]; then
    rm "$POWERPOINT_WEF_DIR/$MANIFEST_FILENAME"
    echo -e "  \033[32m✓ Removed PowerPoint sideload registration\033[0m"
else
    if remove_matching_manifests "$POWERPOINT_WEF_DIR" "$TARGET_MANIFEST_ID"; then
        echo -e "  \033[32m✓ Removed PowerPoint sideload registration\033[0m"
    else
        echo -e "  \033[90m• No OpenCode registration found for PowerPoint\033[0m"
    fi
fi

if [ -f "$EXCEL_WEF_DIR/$MANIFEST_FILENAME" ]; then
    rm "$EXCEL_WEF_DIR/$MANIFEST_FILENAME"
    echo -e "  \033[32m✓ Removed Excel sideload registration\033[0m"
else
    if remove_matching_manifests "$EXCEL_WEF_DIR" "$TARGET_MANIFEST_ID"; then
        echo -e "  \033[32m✓ Removed Excel sideload registration\033[0m"
    else
        echo -e "  \033[90m• No OpenCode registration found for Excel\033[0m"
    fi
fi

echo ""
echo -e "\033[36mSideload cleanup complete.\033[0m"
if [ -d "$APP_PATH" ]; then
    remove_packaged_certificate "$(packaged_user_data_dir)/certs"
    echo "Removed the generated localhost certificate and packaged cert files."
else
    echo "The localhost certificate stays in the system keychain."
    echo "If you want to remove it, open Keychain Access and search for 'localhost'."
fi
echo ""
echo -e "\033[90mTo register again later, run: ./register.sh\033[0m"
