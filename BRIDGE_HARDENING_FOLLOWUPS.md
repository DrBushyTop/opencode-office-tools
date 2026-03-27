# Bridge Hardening Follow-ups

This note captures the remaining bridge-security and local-surface hardening work that we intentionally deferred.

## Current status

- The local HTTP bridge now exposes only the bridge execution surface, not the full `/api` app.
- Bridge execution now requires a bridge token.
- Office host registration/poll/respond now uses a session token plus executor identity.
- Startup regressions caused by eager Word global access were fixed separately.

## Remaining medium-risk items

### 1. Secure-request detection trusts `x-forwarded-proto`

Location: `src/server/api.js`

- `isSecureRequest()` currently accepts `x-forwarded-proto=https` directly.
- If this app ever runs behind a misconfigured proxy, a plain HTTP caller could spoof the header and mint bridge sessions.

Recommended fix:
- Only trust forwarded headers when `app.set("trust proxy", ...)` is explicitly configured.
- Otherwise rely on `req.secure` / TLS termination controlled by Express.

### 2. Token file path still needs stronger local-hardening guarantees

Locations: `src/server.js`, `src/server-prod.js`, `src/server/bridgeTokenPath.js`, `.opencode/lib/office-tool.ts`

- Token files now live in a per-user temp subdirectory with restricted modes.
- We still do not verify ownership, symlink status, or tampering of pre-existing directories/files on read or write.

Recommended fix:
- Verify the token directory is owned by the current user.
- Reject symlinks for both directory and token file.
- Consider atomic create/write with temp file + rename.
- Validate file mode on read before trusting the token.

### 3. Bridge token is still copied into process environment

Locations: `src/server.js`, `src/server-prod.js`

- `process.env.OPENCODE_OFFICE_BRIDGE_TOKEN` is still set for convenience/fallback.
- That broadens exposure to child processes, crash dumps, and future integrations.

Recommended fix:
- Remove env propagation and rely on the token file or a tighter runtime handoff.
- If env fallback remains, scope it to spawned runtime processes only.

### 4. Route-level hardening should get explicit tests

Current tests cover bridge object behavior, but not full HTTP routing behavior.

Recommended tests:
- HTTP bridge cannot access `/api/opencode/*`
- HTTP bridge cannot mint `/api/office-tools/session`
- HTTPS app can mint bridge sessions
- invalid bridge token on `/api/office-tools/execute` returns `401`

### 5. Other pre-existing local-surface issues

Locations: `src/server/api.js`

- `/upload-image` still writes a user-provided filename into a filesystem path.
- `/fetch` is still an unauthenticated arbitrary fetch surface.

Recommended follow-up:
- sanitize uploaded filenames or ignore caller-provided names entirely
- restrict `/fetch` to explicit allowlists or remove it if not required

## Suggested order

1. fix secure-request detection
2. harden token-file ownership/symlink handling
3. remove env-token exposure
4. add route-level regression tests
5. tighten `/upload-image` and `/fetch`
