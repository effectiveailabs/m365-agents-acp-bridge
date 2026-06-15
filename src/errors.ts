export type BridgeErrorCode =
  | 'MS_AUTH_REQUIRED'
  | 'MS_AUTH_EXPIRED'
  | 'MS_TOKEN_SOURCE_UNSUPPORTED'
  | 'MS_TENANT_NOT_CONFIGURED'
  | 'MS_AGENT_NOT_CONFIGURED'
  | 'MS_AGENT_NOT_FOUND'
  | 'MS_INVOKE_FORBIDDEN'
  | 'MS_INVOKE_FAILED'
  | 'MS_RATE_LIMITED'
  | 'MS_UNSUPPORTED_ACTIVITY'
  | 'MS_STREAM_INTERRUPTED'
  | 'MS_SESSION_NOT_FOUND'
  | 'MS_CANCEL_BEST_EFFORT'
  | 'CONFIG_INVALID'
  | 'SECRET_REF_UNSUPPORTED';

export class BridgeError extends Error {
  constructor(
    public readonly code: BridgeErrorCode,
    message: string,
    public readonly status = 500,
    public readonly data?: Record<string, unknown>,
  ) {
    super(message);
    this.name = 'BridgeError';
  }
}

export function asBridgeError(error: unknown, fallbackCode: BridgeErrorCode): BridgeError {
  if (error instanceof BridgeError) {
    return error;
  }

  if (error instanceof Error) {
    return new BridgeError(fallbackCode, error.message);
  }

  return new BridgeError(fallbackCode, 'Unknown bridge error');
}
