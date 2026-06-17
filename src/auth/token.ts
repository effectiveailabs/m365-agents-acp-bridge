export interface DecodedAccessToken {
  aud?: string;
  azp?: string;
  appid?: string;
  exp?: number;
  oid?: string;
  roles?: string[];
  scp?: string;
  tid?: string;
  upn?: string;
}

export function accessTokenFromEnv(
  env: NodeJS.ProcessEnv = process.env,
  tokenEnvName?: string,
): string | undefined {
  if (tokenEnvName) {
    return env[tokenEnvName];
  }

  return env.MICROSOFT_ACCESS_TOKEN ?? env.M365_ACP_MICROSOFT_ACCESS_TOKEN;
}

export function decodeAccessToken(token: string): DecodedAccessToken {
  const parts = token.split('.');
  if (parts.length < 2 || !parts[1]) {
    throw new Error('Access token is not a JWT');
  }

  const payload = JSON.parse(Buffer.from(parts[1], 'base64url').toString('utf8')) as Record<
    string,
    unknown
  >;

  return {
    aud: stringClaim(payload.aud),
    azp: stringClaim(payload.azp),
    appid: stringClaim(payload.appid),
    exp: numberClaim(payload.exp),
    oid: stringClaim(payload.oid),
    roles: stringArrayClaim(payload.roles),
    scp: stringClaim(payload.scp),
    tid: stringClaim(payload.tid),
    upn: stringClaim(payload.upn),
  };
}

export function hasScope(token: DecodedAccessToken, scope: string): boolean {
  return token.scp?.split(/\s+/).includes(scope) ?? false;
}

function stringClaim(value: unknown): string | undefined {
  return typeof value === 'string' && value.length > 0 ? value : undefined;
}

function numberClaim(value: unknown): number | undefined {
  return typeof value === 'number' && Number.isFinite(value) ? value : undefined;
}

function stringArrayClaim(value: unknown): string[] | undefined {
  if (!Array.isArray(value)) {
    return undefined;
  }

  const strings = value.filter((item): item is string => typeof item === 'string');
  return strings.length > 0 ? strings : undefined;
}
