# OpenCode Office Add-in

A Microsoft Office add-in that brings OpenCode into Word, Excel, PowerPoint, and OneNote.

> **Note:** This project is not affiliated with, endorsed by, or associated with [anomalyco/opencode](https://github.com/anomalyco/opencode/tree/dev) in any way. It is an independent project built on the opencode SDK.

The add-in starts or attaches to an OpenCode runtime locally, injects the bundled Office tool set from `.opencode/`, and routes tool execution back into the active Office task pane.

> **Requires [OpenCode CLI](https://opencode.ai/docs) installed on your system.** The add-in spawns a local `opencode serve` process under the hood. Install it with `brew install anomalyco/tap/opencode`, `npm i -g opencode-ai@latest`, or any method listed in the [getting started guide](GETTING_STARTED.md#installing-opencode).

## Getting Started

**👉 See [GETTING_STARTED.md](GETTING_STARTED.md) for setup instructions.**

**📖 See [TOOLS_CATALOG.md](TOOLS_CATALOG.md) for available OpenCode tools.**

The add-in bundles its own OpenCode tool configuration, so users do not need to create a separate `.opencode` setup to access the Office tools.

The getting started guide walks you through running the add-in locally using the tray app. For packaged desktop installers, see [installer/README.md](installer/README.md).

## Flow Overview

```text
+-------------------+
| Office Add-in UI  |
| Word / Excel /    |
| PowerPoint /      |
| OneNote           |
+---------+---------+
          |
          | user prompt / action
          v
+-------------------+
| Office bridge     |
| task pane + local |
| add-in server     |
+---------+---------+
          |
          | starts or attaches
          v
+-------------------+
| OpenCode runtime  |
| agent + tool      |
| orchestration     |
+---------+---------+
          |
          | calls bundled tools
          v
+-------------------+
| .opencode tools   |
| Office tools +    |
| local utilities   |
+----+---------+----+
     |         |
     |         +------------------------+
     |                                  |
     v                                  v
+------------+                   +--------------+
| Read state |                   | Make edits   |
| document,  |                   | update Word, |
| slides,    |                   | Excel, PPT,  |
| workbook   |                   | OneNote      |
+-----+------+                   +------+-------+
      |                                   |
      +----------------+------------------+
                       |
                       | results / updated state
                       v
                +-------------+
                | Office app  |
                | reflects    |
                | changes     |
                +-------------+
```

## Project Structure

```
├── src/
│   ├── server.js          # Dev wrapper that mounts Vite on the shared HTTPS runtime
│   ├── server-prod.js     # Production wrapper that serves dist/ on the shared HTTPS runtime
│   ├── server/            # Shared HTTP runtime, OpenCode runtime, API routes, and Office bridge adapters
│   ├── shared/            # Shared Office metadata and generated tool registry
│   ├── tray/              # Electron tray app entrypoint
│   └── ui/                # React task pane frontend and Office tools
├── .opencode/             # Bundled OpenCode tools and config used by the add-in
├── assets/                # Tray icons and packaged app assets
├── dist/                  # Built frontend assets
├── certs/                 # Development localhost certs and OpenSSL config
├── scripts/               # Build and packaging helper scripts
├── manifest.xml           # Office add-in manifest for Word, Excel, PowerPoint, and OneNote
├── installer/             # Installer resources (Electron Builder)
│   ├── macos/             # macOS post-install scripts
│   └── windows/           # Windows NSIS scripts
├── register.sh/.ps1       # Setup scripts (trust cert, register manifest)
└── unregister.sh/.ps1     # Cleanup scripts
```

## Scripts

| Command | Description |
|---------|-------------|
| `bun run dev` | Start development server with hot reload |
| `bun run start` | Run the built frontend on the shared production server |
| `bun run start:tray` | Build the app and launch the local Electron tray runtime |
| `bun run build` | Build frontend for production |
| `bun run test` | Regenerate Office tool metadata and run the Vitest suite |
| `bun run build:installer` | Build installer for current platform |
| `bun run build:installer:mac` | Build macOS .pkg installer |
| `bun run build:installer:win` | Build Windows .exe installer |

## Unregistering Add-in

```bash
./unregister.sh      # macOS
.\unregister.ps1     # Windows
```

## Troubleshooting

### Add-in not appearing
1. Ensure the local service is running: visit https://localhost:52390
2. Look for the OpenCode icon in the system tray (Windows) or menu bar (macOS)
3. Restart the Office application
4. Clear Office cache and try again

### Runtime mode
- The add-in prefers an attached OpenCode runtime when `OPENCODE_OFFICE_RUNTIME_URL` or `OPENCODE_RUNTIME_URL` is configured and reachable.
- Otherwise it spawns a local OpenCode runtime using the bundled `.opencode` config.

### SSL Certificate errors
1. Re-run the register script or installer
2. Development mode uses `certs/localhost.pem`
3. Installed builds generate and trust a per-user localhost certificate automatically

### Service not starting after install
- **Windows**: Check `HKCU\Software\Microsoft\Windows\CurrentVersion\Run` for `OpenCodeOfficeAddin`
- **macOS**: Run `launchctl list | grep com.opencode.office-addin`
- Launch the installed app manually once and use the tray/menu bar action to open the debug log if needed
