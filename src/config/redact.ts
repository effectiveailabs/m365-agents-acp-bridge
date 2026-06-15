const SENSITIVE_KEY_PATTERN =
  /(access[_-]?token|refresh[_-]?token|authorization|client[_-]?secret|connection[_-]?string|direct[_-]?connect[_-]?url|password|secret)/i;

export function redactSensitive(value: unknown): unknown {
  return redactValue(value, undefined);
}

export function redactString(value: string): string {
  if (value.length <= 8) {
    return '[redacted]';
  }

  return `${value.slice(0, 4)}...[redacted]`;
}

function redactValue(value: unknown, key: string | undefined): unknown {
  if (value === null || value === undefined) {
    return value;
  }

  if (typeof value === 'string') {
    return key && SENSITIVE_KEY_PATTERN.test(key) ? redactString(value) : value;
  }

  if (Array.isArray(value)) {
    return value.map((item) => redactValue(item, key));
  }

  if (typeof value === 'object') {
    const output: Record<string, unknown> = {};
    for (const [entryKey, entryValue] of Object.entries(value)) {
      if (entryKey === 'secretRef' && typeof entryValue === 'string') {
        output[entryKey] = entryValue;
        continue;
      }

      output[entryKey] = redactValue(entryValue, entryKey);
    }
    return output;
  }

  return value;
}
