import { getCopilotStudioConnectionUrl } from '@microsoft/agents-copilotstudio-client';
import type { BridgeConfig } from '../config/types.js';
import { decodeAccessToken, hasScope } from '../auth/token.js';
import { connectionSettingsForCopilot } from '../microsoft/connectionSettings.js';

const REQUIRED_SCOPE = 'CopilotStudio.Copilots.Invoke';
const POWER_PLATFORM_AUDIENCE = 'https://api.powerplatform.com';

export type DoctorSeverity = 'ok' | 'warn' | 'fail';

export interface DoctorCheck {
  severity: DoctorSeverity;
  name: string;
  message: string;
}

export interface DoctorResult {
  ok: boolean;
  checks: DoctorCheck[];
}

export interface DoctorOptions {
  config: BridgeConfig;
  accessToken?: string;
  env?: NodeJS.ProcessEnv;
  now?: Date;
}

export async function runDoctor(options: DoctorOptions): Promise<DoctorResult> {
  const checks: DoctorCheck[] = [];
  const env = options.env ?? process.env;

  add(checks, 'ok', 'config', `Loaded ${options.config.agents.length} agent config(s)`);

  if (options.config.auth.mode === 'external_token') {
    add(checks, 'ok', 'auth mode', 'external_token enabled');
  } else if (options.config.auth.mode === 'interactive') {
    add(checks, 'warn', 'auth mode', 'interactive mode is for local development only');
  } else {
    add(checks, 'fail', 'auth mode', `${options.config.auth.mode} is not supported in v1`);
  }

  for (const agent of options.config.agents) {
    await checkAgentTarget(
      checks,
      agent.id,
      options.config.agents.length,
      agent.copilotStudio,
      env,
    );
  }

  checkToken(checks, options.accessToken, options.now ?? new Date());

  return {
    ok: !checks.some((check) => check.severity === 'fail'),
    checks,
  };
}

export function formatDoctorResult(result: DoctorResult): string {
  const lines = ['M365 ACP doctor'];
  for (const check of result.checks) {
    lines.push(`[${check.severity}] ${check.name}: ${check.message}`);
  }
  lines.push(result.ok ? 'Result: ok' : 'Result: failed');
  return `${lines.join('\n')}\n`;
}

async function checkAgentTarget(
  checks: DoctorCheck[],
  agentId: string,
  agentCount: number,
  copilot: BridgeConfig['agents'][number]['copilotStudio'],
  env: NodeJS.ProcessEnv,
): Promise<void> {
  const prefix = agentCount > 1 ? `agent ${agentId}` : 'agent';
  const hasDirect = Boolean(copilot.directConnectUrl || copilot.connectionString);

  if (hasDirect) {
    add(checks, 'ok', `${prefix} config`, 'uses direct connect URL/connection string secret ref');
  } else {
    add(
      checks,
      'ok',
      `${prefix} config`,
      `uses expanded config environment=${copilot.environmentId}, schema=${
        copilot.schemaName ?? copilot.agentIdentifier
      }`,
    );
  }

  try {
    const settings = await connectionSettingsForCopilot(copilot, env);
    const url = getCopilotStudioConnectionUrl(settings);
    add(checks, 'ok', `${prefix} endpoint`, `Microsoft SDK target resolves to ${hostFor(url)}`);
  } catch (error) {
    add(checks, 'fail', `${prefix} endpoint`, errorMessage(error));
  }
}

function checkToken(checks: DoctorCheck[], accessToken: string | undefined, now: Date): void {
  if (!accessToken) {
    add(
      checks,
      'fail',
      'token',
      'No Microsoft access token found. Set MICROSOFT_ACCESS_TOKEN or M365_ACP_MICROSOFT_ACCESS_TOKEN.',
    );
    return;
  }

  let decoded: ReturnType<typeof decodeAccessToken>;
  try {
    decoded = decodeAccessToken(accessToken);
  } catch (error) {
    add(checks, 'fail', 'token', errorMessage(error));
    return;
  }

  if (decoded.aud === POWER_PLATFORM_AUDIENCE) {
    add(checks, 'ok', 'token audience', decoded.aud);
  } else {
    add(
      checks,
      'fail',
      'token audience',
      `Expected ${POWER_PLATFORM_AUDIENCE}, got ${decoded.aud ?? 'missing'}`,
    );
  }

  if (hasScope(decoded, REQUIRED_SCOPE)) {
    add(checks, 'ok', 'token scope', REQUIRED_SCOPE);
  } else if (decoded.roles?.includes(REQUIRED_SCOPE)) {
    add(
      checks,
      'fail',
      'token scope',
      `Found app-only roles:${REQUIRED_SCOPE}; v1 requires delegated scp:${REQUIRED_SCOPE}`,
    );
  } else {
    add(
      checks,
      'fail',
      'token scope',
      `Missing delegated scope ${REQUIRED_SCOPE}; token scp=${decoded.scp ?? 'missing'}`,
    );
  }

  if (!decoded.exp) {
    add(checks, 'warn', 'token expiry', 'Token has no exp claim');
  } else {
    const secondsRemaining = decoded.exp - Math.floor(now.getTime() / 1000);
    if (secondsRemaining <= 0) {
      add(checks, 'fail', 'token expiry', 'Token is expired');
    } else if (secondsRemaining < 300) {
      add(checks, 'warn', 'token expiry', `Token expires in ${secondsRemaining}s`);
    } else {
      add(checks, 'ok', 'token expiry', `Token expires in ${secondsRemaining}s`);
    }
  }

  if (decoded.upn || decoded.oid) {
    add(checks, 'ok', 'token identity', decoded.upn ? 'upn present' : 'oid present');
  } else {
    add(checks, 'warn', 'token identity', 'No upn or oid claim found');
  }
}

function add(checks: DoctorCheck[], severity: DoctorSeverity, name: string, message: string): void {
  checks.push({ severity, name, message });
}

function hostFor(rawUrl: string): string {
  try {
    return new URL(rawUrl).host;
  } catch {
    return 'configured endpoint';
  }
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
