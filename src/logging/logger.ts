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
      writeLog('debug', message, fields);
    }
  },
  info(message, fields) {
    writeLog('info', message, fields);
  },
  warn(message, fields) {
    writeLog('warn', message, fields);
  },
  error(message, fields) {
    writeLog('error', message, fields);
  },
};

function writeLog(level: string, message: string, fields?: Record<string, unknown>): void {
  const payload = fields ? ` ${JSON.stringify(redactSensitive(fields))}` : '';
  const line = `[${level}] ${message}${payload}`;
  if (level === 'error' || level === 'warn') {
    console.error(line);
    return;
  }
  console.log(line);
}
