#!/bin/bash

# Resolve script-relative resources once so the script works in both repo and packaged layouts.
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
APP_PATH="$SCRIPT_DIR/OpenCode Office Add-in.app"
MANIFEST_FILENAME="manifest.xml"

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

# Prefer packaged resources when the app bundle is present; otherwise use repo-local assets.
if [ -d "$APP_PATH" ]; then
    MANIFEST_PATH="$APP_PATH/Contents/Resources/manifest.xml"
    CERT_PATH="$APP_PATH/Contents/Resources/certs/localhost.pem"
else
    MANIFEST_PATH="$SCRIPT_DIR/manifest.xml"
    CERT_PATH="$SCRIPT_DIR/certs/localhost.pem"
fi

TARGET_MANIFEST_ID="$(manifest_id "$MANIFEST_PATH")"

echo -e "\033[36mPreparing OpenCode Office Add-in on macOS...\033[0m"
echo ""

# Step 0: Clear Gatekeeper quarantine flags from downloaded app bundles.
if [ -d "$APP_PATH" ]; then
    echo -e "\033[33mStep 0: Clearing quarantine flags from the app bundle...\033[0m"
    xattr -cr "$APP_PATH" 2>/dev/null
    echo -e "  \033[32m✓ App bundle is ready to launch\033[0m"
    echo ""
fi

# Step 1: Trust the localhost HTTPS certificate used by the local add-in server.
echo -e "\033[33mStep 1: Trusting the localhost HTTPS certificate...\033[0m"

if [ ! -f "$CERT_PATH" ]; then
    echo -e "\033[31mError: Missing certificate at $CERT_PATH\033[0m"
    echo -e "\033[31mThe local HTTPS endpoint cannot start without it.\033[0m"
    exit 1
fi

# Reuse trust if the exact certificate fingerprint is already installed.
if security find-certificate -c "localhost" -a -Z | grep -q "$(openssl x509 -in "$CERT_PATH" -fingerprint -noout | cut -d= -f2)"; then
    echo -e "  \033[32m✓ Certificate trust is already in place\033[0m"
else
    if sudo security add-trusted-cert -d -r trustRoot -k /Library/Keychains/System.keychain "$CERT_PATH"; then
        echo -e "  \033[32m✓ Certificate trust installed\033[0m"
    else
        echo -e "  \033[31mError: Could not add the certificate to the system keychain.\033[0m"
        echo -e "  \033[31mRun this script in an interactive terminal and approve the sudo prompt.\033[0m"
        exit 1
    fi
fi

echo ""

# Step 2: Refresh the Office sideload manifest for each supported host.
echo -e "\033[33mStep 2: Refreshing Office sideload manifest...\033[0m"
echo "  Manifest source: $MANIFEST_PATH"

# Ensure each host-specific sideload folder exists.
WORD_WEF_DIR="$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef"
POWERPOINT_WEF_DIR="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
EXCEL_WEF_DIR="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
ONENOTE_WEF_DIR="$HOME/Library/Containers/com.microsoft.onenote.mac/Data/Documents/wef"

mkdir -p "$WORD_WEF_DIR"
mkdir -p "$POWERPOINT_WEF_DIR"
mkdir -p "$EXCEL_WEF_DIR"
mkdir -p "$ONENOTE_WEF_DIR"

# Replace older OpenCode registrations before copying the active manifest into each host folder.
for WEF_DIR in "$WORD_WEF_DIR" "$POWERPOINT_WEF_DIR" "$EXCEL_WEF_DIR" "$ONENOTE_WEF_DIR"; do
    remove_matching_manifests "$WEF_DIR" "$TARGET_MANIFEST_ID"
    cp "$MANIFEST_PATH" "$WEF_DIR/$MANIFEST_FILENAME"
done

echo -e "  \033[32m✓ Word sideload registration updated\033[0m"
echo -e "  \033[32m✓ PowerPoint sideload registration updated\033[0m"
echo -e "  \033[32m✓ Excel sideload registration updated\033[0m"
echo -e "  \033[32m✓ OneNote sideload registration updated\033[0m"
echo ""

echo -e "\033[36mSetup complete. Next steps:\033[0m"
echo "1. Close Word, PowerPoint, Excel, and OneNote if they are open"
echo "2. Launch the tray runtime: bun run start:tray"
echo "3. Open Word, PowerPoint, Excel, or OneNote"
echo "4. Go to Insert > Add-ins > My Add-ins and look for 'OpenCode'"
echo ""
echo -e "\033[90mTo remove the sideload registration later, run: ./unregister.sh\033[0m"
