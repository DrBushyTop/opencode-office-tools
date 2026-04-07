# Getting Started

Run the OpenCode Office Add-in from this package.

## Prerequisites

| Software | Download |
|----------|----------|
| **Microsoft Office** | Word, PowerPoint, or Excel (Microsoft 365 or Office 2019+) |

## Setup

### 1. Install the Package

**macOS:** Run the `OpenCodeOfficeAddin-<version>.pkg` installer

**Windows:** Run `OpenCode Office Add-in Setup.exe`

The installer generates and trusts a localhost certificate for this machine and refreshes the Office sideload registration.

### 2. Launch the App

If the app does not start automatically:

**macOS:** Open `OpenCode Office Add-in.app` from Applications

**Windows:** Run `OpenCode Office Add-in.exe`

You should see the OpenCode icon appear in your system tray (Windows) or menu bar (macOS).

## Adding the Add-in in Office

1. Confirm you see the OpenCode service running in your system tray/menu bar.

2. **Open** Word, PowerPoint, or Excel
   > **Close and reopen the app if it was already running before installation completed**

3. Go to **Insert** → **Add-ins** → **My Add-ins**

4. Look for the **OpenCode** add-in. Write text or paste images to get started.

## Troubleshooting

### Add-in not showing up?
- Make sure the tray app is running (check for the icon in your system tray/menu bar)
- Completely quit and restart the Office application
- Re-run the installer

### SSL Certificate errors?
- Re-run the installer
- On macOS, you may need to enter your password to trust the certificate during install
