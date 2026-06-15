import { readFile } from 'node:fs/promises';
import type { SecretRef } from './types.js';

export class SecretRefError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'SecretRefError';
  }
}

export async function resolveSecretRef(
  secret: SecretRef,
  env: NodeJS.ProcessEnv = process.env,
): Promise<string> {
  const ref = secret.secretRef;

  if (ref.startsWith('env:')) {
    const name = ref.slice('env:'.length);
    const value = env[name];
    if (!value) {
      throw new SecretRefError(`Environment secret ref ${name} is not set`);
    }
    return value;
  }

  if (ref.startsWith('file:')) {
    const path = ref.slice('file:'.length);
    if (!path.startsWith('/')) {
      throw new SecretRefError('file: secret refs must use absolute paths');
    }
    return (await readFile(path, 'utf8')).trim();
  }

  throw new SecretRefError(`Unsupported secret ref scheme: ${ref}`);
}
