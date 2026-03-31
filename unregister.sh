#!/bin/bash

# Get the directory where the script is located
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
APP_PATH="$SCRIPT_DIR/OpenCode Office Add-in.app"
MANIFEST_FILENAME="manifest.xml"

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

echo -e "\033[36mUnregistering Office Add-in from macOS...\033[0m"
echo ""

# Define directories
WORD_WEF_DIR="$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef"
POWERPOINT_WEF_DIR="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
EXCEL_WEF_DIR="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
ONENOTE_WEF_DIR="$HOME/Library/Containers/com.microsoft.onenote.mac/Data/Documents/wef"

# Remove manifest from Word directory
if [ -f "$WORD_WEF_DIR/$MANIFEST_FILENAME" ]; then
    rm "$WORD_WEF_DIR/$MANIFEST_FILENAME"
    echo -e "  \033[32m✓ Removed add-in from Word\033[0m"
else
    if remove_matching_manifests "$WORD_WEF_DIR" "$TARGET_MANIFEST_ID"; then
        echo -e "  \033[32m✓ Removed add-in from Word\033[0m"
    else
        echo -e "  \033[90m• Add-in not found in Word directory\033[0m"
    fi
fi

# Remove manifest from PowerPoint directory
if [ -f "$POWERPOINT_WEF_DIR/$MANIFEST_FILENAME" ]; then
    rm "$POWERPOINT_WEF_DIR/$MANIFEST_FILENAME"
    echo -e "  \033[32m✓ Removed add-in from PowerPoint\033[0m"
else
    if remove_matching_manifests "$POWERPOINT_WEF_DIR" "$TARGET_MANIFEST_ID"; then
        echo -e "  \033[32m✓ Removed add-in from PowerPoint\033[0m"
    else
        echo -e "  \033[90m• Add-in not found in PowerPoint directory\033[0m"
    fi
fi

# Remove manifest from Excel directory
if [ -f "$EXCEL_WEF_DIR/$MANIFEST_FILENAME" ]; then
    rm "$EXCEL_WEF_DIR/$MANIFEST_FILENAME"
    echo -e "  \033[32m✓ Removed add-in from Excel\033[0m"
else
    if remove_matching_manifests "$EXCEL_WEF_DIR" "$TARGET_MANIFEST_ID"; then
        echo -e "  \033[32m✓ Removed add-in from Excel\033[0m"
    else
        echo -e "  \033[90m• Add-in not found in Excel directory\033[0m"
    fi
fi

# Remove manifest from OneNote directory
if [ -f "$ONENOTE_WEF_DIR/$MANIFEST_FILENAME" ]; then
    rm "$ONENOTE_WEF_DIR/$MANIFEST_FILENAME"
    echo -e "  \033[32m✓ Removed add-in from OneNote\033[0m"
else
    if remove_matching_manifests "$ONENOTE_WEF_DIR" "$TARGET_MANIFEST_ID"; then
        echo -e "  \033[32m✓ Removed add-in from OneNote\033[0m"
    else
        echo -e "  \033[90m• Add-in not found in OneNote directory\033[0m"
    fi
fi

echo ""
echo -e "\033[36mUnregistration complete!\033[0m"
echo "Note: The SSL certificate remains in the system keychain."
echo "To remove it, use Keychain Access app and search for 'localhost'."
echo ""
echo -e "\033[90mTo re-register, run: ./register.sh\033[0m"
