import { redactSensitive } from '../config/redact.js';

export interface Logger {
  debug(message: string, fields?: Record<string, unknown>): void;
  info(message: string, fields?: Record<string, unknown>): void;
  warn(message: string, fields?: Record<string, unknown>): void;
  error(message: string, fields?: Record<string, unknown>): void;
}

export const consoleLogger: Logger = {
  debug(message, fields) {
    if (process.env.M365_ACP_LOG_LEVEL === 'debug') {
      writeLog(console.log, 'debug', message, fields);
    }
  },
  info(message, fields) {
    writeLog(console.log, 'info', message, fields);
  },
  warn(message, fields) {
    writeLog(console.error, 'warn', message, fields);
  },
  error(message, fields) {
    writeLog(console.error, 'error', message, fields);
  },
};

export const stderrLogger: Logger = {
  debug(message, fields) {
    if (process.env.M365_ACP_LOG_LEVEL === 'debug') {
      writeLog(console.error, 'debug', message, fields);
    }
  },
  info(message, fields) {
    writeLog(console.error, 'info', message, fields);
  },
  warn(message, fields) {
    writeLog(console.error, 'warn', message, fields);
  },
  error(message, fields) {
    writeLog(console.error, 'error', message, fields);
  },
};

function writeLog(
  write: (line: string) => void,
  level: string,
  message: string,
  fields?: Record<string, unknown>,
): void {
  const payload = fields ? ` ${JSON.stringify(redactSensitive(fields))}` : '';
  const line = `[${level}] ${message}${payload}`;
  write(line);
}
