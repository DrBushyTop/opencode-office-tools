const fs = require('fs');
const path = require('path');
const util = require('util');

function getLogFilePath() {
  const configured = process.env.OPENCODE_OFFICE_LOG_FILE;
  if (configured) return configured;
  return path.resolve(process.cwd(), '.opencode', 'debug.log');
}

function ensureLogDirectory() {
  fs.mkdirSync(path.dirname(getLogFilePath()), { recursive: true });
}

function normalizeDetail(detail) {
  if (detail == null) return '';
  if (detail instanceof Error) {
    return detail.stack || detail.message;
  }
  if (typeof detail === 'string') return detail;
  return util.inspect(detail, { depth: 6, breakLength: 120, maxArrayLength: 20 });
}

function writeLog(level, scope, message, detail) {
  const timestamp = new Date().toISOString();
  const suffix = normalizeDetail(detail);
  const line = `${timestamp} ${level.toUpperCase()} [${scope}] ${message}${suffix ? `\n${suffix}` : ''}\n`;

  try {
    ensureLogDirectory();
    fs.appendFileSync(getLogFilePath(), line, 'utf8');
  } catch (error) {
    const fallback = `devLogger failed: ${error instanceof Error ? error.message : String(error)}`;
    process.stderr.write(`${fallback}\n`);
    process.stderr.write(line);
  }
}

function logInfo(scope, message, detail) {
  writeLog('info', scope, message, detail);
}

function logWarn(scope, message, detail) {
  writeLog('warn', scope, message, detail);
}

function logError(scope, message, detail) {
  writeLog('error', scope, message, detail);
}

module.exports = {
  getLogFilePath,
  logInfo,
  logWarn,
  logError,
};
