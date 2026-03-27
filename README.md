# OpenCode Office Add-in

A Microsoft Office add-in that brings OpenCode into Word, Excel, and PowerPoint.

The add-in starts or attaches to an OpenCode runtime locally, injects the bundled Office tool set from `.opencode/`, and routes tool execution back into the active Office task pane.

## Getting Started

**👉 See [GETTING_STARTED.md](GETTING_STARTED.md) for setup instructions.**

**📖 See [TOOLS_CATALOG.md](TOOLS_CATALOG.md) for available OpenCode tools.**

The add-in bundles its own OpenCode tool configuration, so users do not need to create a separate `.opencode` setup to access the Office tools.

The getting started guide walks you through running the add-in locally using the tray app. Standalone installers are in development and will be available once code signing is complete.

## Flow Overview

```text
+-------------------+
| Office Add-in UI  |
| Word / Excel /    |
| PowerPoint        |
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
| slides,    |                   | Excel, PPT   |
| workbook   |                   | content      |
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

## Office Videos

### PowerPoint

https://github.com/user-attachments/assets/4c2731e4-e157-4968-842f-e496a6e8ed8b

### Excel


https://github.com/user-attachments/assets/42478d69-fd26-415e-8ef7-4efe8450d695

### Word

https://github.com/user-attachments/assets/41408f8d-a9b8-45b6-a826-f50931c7c249

## Project Structure

```
├── src/
│   ├── server.js          # Dev server (Vite + Express)
│   ├── server-prod.js     # Production server (static files)
│   ├── server/            # OpenCode runtime and Office bridge adapters
│   └── ui/                # React frontend
├── .opencode/             # Bundled OpenCode tools and config used by the add-in
├── dist/                  # Built frontend assets
├── certs/                 # SSL certificates for localhost
├── manifest.xml           # Office add-in manifest
├── installer/             # Installer resources (Electron Builder)
│   ├── macos/             # macOS post-install scripts
│   └── windows/           # Windows NSIS scripts
├── register.sh/.ps1       # Setup scripts (trust cert, register manifest)
└── unregister.sh/.ps1     # Cleanup scripts
```

## Scripts

| Command | Description |
|---------|-------------|
| `npm run dev` | Start development server with hot reload |
| `npm run start` | Run production server standalone |
| `npm run start:tray` | Run Electron tray app locally |
| `npm run build` | Build frontend for production |
| `npm test` | Run focused Vitest coverage for history, permissions, and tool exposure |
| `npm run build:installer` | Build installer for current platform |
| `npm run build:installer:mac` | Build macOS .dmg installer |
| `npm run build:installer:win` | Build Windows .exe installer |

## Unregistering Add-in

```bash
./unregister.sh      # macOS
.\unregister.ps1     # Windows
```

## Troubleshooting

### Add-in not appearing
1. Ensure the server is running: visit https://localhost:52390
2. Look for the OpenCode icon in the system tray (Windows) or menu bar (macOS)
3. Restart the Office application
4. Clear Office cache and try again

### Runtime mode
- The add-in prefers an attached OpenCode runtime when `OPENCODE_OFFICE_RUNTIME_URL` or `OPENCODE_RUNTIME_URL` is configured and reachable.
- Otherwise it spawns a local OpenCode runtime using the bundled `.opencode` config.

### SSL Certificate errors
1. Re-run the register script or installer
2. Or manually trust `certs/localhost.pem`

### Service not starting after install
- **Windows**: Check Task Scheduler or startup entries for "OpenCodeOfficeAddin"
- **macOS**: Run `launchctl list | grep opencode` and check `/tmp/opencode-office-addin.log`
