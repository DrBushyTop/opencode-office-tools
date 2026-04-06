# Getting Started (Local Development)

Run the OpenCode Office Add-in locally using the tray app—no installers required.

The add-in injects its bundled Office tools into the OpenCode runtime automatically, so you do not need to create your own `.opencode` folder first.

> **📖 See [TOOLS_CATALOG.md](TOOLS_CATALOG.md) for a complete list of available OpenCode tools for Word, PowerPoint, and Excel.**

## Prerequisites

Install the following software:

| Software | Download |
|----------|----------|
| **OpenCode CLI** | See [installation options](#installing-opencode) below |
| **Node.js 20+** | [nodejs.org](https://nodejs.org/) |
| **Git** | [git-scm.com](https://git-scm.com/downloads) |
| **Microsoft Office** | Word, PowerPoint, or Excel (Microsoft 365 or Office 2019+) |

### Installing OpenCode

The add-in requires the **`opencode` CLI** to be installed on your system. When the add-in starts, it spawns a local OpenCode server process using the `opencode` binary. If the binary is not found, the add-in will fail to start a session.

Install OpenCode using any of these methods:

```bash
# Quick install
curl -fsSL https://opencode.ai/install | bash

# Package managers
npm i -g opencode-ai@latest        # or bun/pnpm/yarn
brew install anomalyco/tap/opencode # macOS and Linux
scoop install opencode              # Windows
choco install opencode              # Windows
```

For the full list of installation options, see the [OpenCode documentation](https://opencode.ai/docs).

The add-in searches for the `opencode` binary in these locations (in addition to your system `PATH`):
- `~/.opencode/bin`
- `~/.local/bin`
- `~/.bun/bin`
- `/opt/homebrew/bin` (macOS)
- `/usr/local/bin`
- `~/AppData/Local/Programs/opencode/bin` (Windows)

> **Tip:** If you already have an OpenCode session running, the add-in can attach to it instead of spawning a new one. Set `OPENCODE_OFFICE_RUNTIME_URL` or `OPENCODE_RUNTIME_URL` to your runtime's URL (see [Setup](#3-start-the-tray-application) below).

## Setup

### 1. Clone and Install Dependencies

```bash
cd /path/to/opencode-officeplugins
bun install
```

Optional but useful checks:

```bash
bun run test
bun run build
```

### 2. Register the Add-in

This trusts the SSL certificate and registers the manifest with Office.

**macOS:**
```bash
./register.sh
```

**Windows (PowerShell as Administrator):**
```powershell
.\register.ps1
```

### 3. Start the Tray Application

```bash
bun run start:tray
```

You should see the OpenCode icon appear in your system tray (Windows) or menu bar (macOS).

For a direct development flow instead of the tray app, run:

```bash
bun run build
bun run dev
```

That starts the local add-in server and points the spawned OpenCode runtime at the bundled `.opencode` config.

If you already have an OpenCode runtime you want the add-in to reuse, set one of these before starting the app:

```bash
export OPENCODE_OFFICE_RUNTIME_URL=http://127.0.0.1:4096
# or
export OPENCODE_RUNTIME_URL=http://127.0.0.1:4096
```

If neither variable is set, the add-in spawns its own local OpenCode runtime automatically.

## Adding the Add-in in Office
1. Confirm you see the OpenCode service running in your macOS or Windows tray.
<img width="211" height="159" alt="image" src="https://github.com/user-attachments/assets/97bd61d2-6977-48e4-bf05-cd1529afa04d" />

2. **Open** Word, PowerPoint, or Excel
3. <img width="203" height="66" alt="image" src="https://github.com/user-attachments/assets/653e8c6f-7e93-447e-ac07-d0c8cf3834dd" />
> **Close and reopen the app if it was already running before registration**

4. Go to **Insert** → **Add-ins** → **My Add-ins**
<img width="459" height="324" alt="image" src="https://github.com/user-attachments/assets/fc157744-a0a0-4975-86d1-380736e2bb12" />

5. Look for the **OpenCode** add-in. Write text or paste images to get started.
<img width="358" height="352" alt="image" src="https://github.com/user-attachments/assets/e06d89a5-5fa8-4940-92b6-e60b04c1e5c7" />

6. Have fun!

https://github.com/user-attachments/assets/5bb771d3-0bf6-4b7b-8e6c-757a085b3131

## Troubleshooting

### Add-in not showing up?
- Make sure the tray app is running (check for the icon in your system tray/menu bar)
- Completely quit and restart the Office application
- Re-run the register script

### SSL Certificate errors?
- Re-run `./register.sh` (macOS) or `.\register.ps1` (Windows)
- On macOS, you may need to enter your password to trust the certificate

### Want to use the dev server with hot reload instead?
```bash
bun run dev
```
This starts the development server on port 52390 with hot reload and a local Office bridge on port 52391.

## Uninstalling

```bash
./unregister.sh      # macOS
.\unregister.ps1     # Windows
```
