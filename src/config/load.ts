import { readFile } from 'node:fs/promises';
import { z } from 'zod';
import { redactSensitive } from './redact.js';
import type { BridgeConfig, LoadConfigOptions } from './types.js';

const secretRefSchema = z.object({
  secretRef: z
    .string()
    .min(1)
    .regex(/^(env|file):/),
});

const copilotStudioSchema = z
  .object({
    tenantId: z.string().min(1).optional(),
    environmentId: z.string().min(1).optional(),
    schemaName: z.string().min(1).optional(),
    agentIdentifier: z.string().min(1).optional(),
    clientId: z.string().min(1).optional(),
    clientSecret: secretRefSchema.optional(),
    connectionString: secretRefSchema.optional(),
    directConnectUrl: secretRefSchema.optional(),
    cloud: z.string().min(1).optional(),
    customPowerPlatformCloud: z.string().min(1).optional(),
    useExperimentalEndpoint: z.boolean().optional(),
  })
  .strict()
  .superRefine((value, ctx) => {
    const hasDirect = Boolean(value.connectionString || value.directConnectUrl);
    const hasExpanded = Boolean(value.environmentId && (value.schemaName || value.agentIdentifier));

    if (!hasDirect && !hasExpanded) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message:
          'Copilot Studio config requires directConnectUrl/connectionString secret ref or expanded environmentId plus schemaName/agentIdentifier',
      });
    }
  });

const configSchema = z
  .object({
    server: z
      .object({
        host: z.string().min(1).default('127.0.0.1'),
        port: z.number().int().min(0).max(65535).default(3838),
      })
      .default({ host: '127.0.0.1', port: 3838 }),
    auth: z
      .object({
        mode: z.enum(['external_token', 'interactive', 'obo']).default('external_token'),
        tokenSource: z
          .enum(['authorization_header', 'authenticate_meta', 'env'])
          .optional()
          .default('authorization_header'),
      })
      .default({ mode: 'external_token', tokenSource: 'authorization_header' }),
    agents: z.array(
      z
        .object({
          id: z.string().min(1),
          displayName: z.string().min(1).optional(),
          description: z.string().min(1).optional(),
          copilotStudio: copilotStudioSchema,
        })
        .strict(),
    ),
    logLevel: z.enum(['debug', 'info', 'warn', 'error']).optional(),
  })
  .strict();

export async function loadBridgeConfig(options: LoadConfigOptions = {}): Promise<BridgeConfig> {
  const env = options.env ?? process.env;
  const fileConfig = options.configFile ? await readJsonConfig(options.configFile) : {};
  const envConfig = configFromEnv(env);
  const merged = mergeConfig(fileConfig, envConfig);
  const config = configSchema.parse(merged);
  validateBridgeConfig(config);
  return config;
}

export function parseBridgeConfig(input: unknown): BridgeConfig {
  const config = configSchema.parse(input);
  validateBridgeConfig(config);
  return config;
}

export function validateBridgeConfig(config: BridgeConfig): void {
  if (config.auth.mode === 'obo') {
    throw new Error('Auth mode "obo" is not supported in v1');
  }

  const seen = new Set<string>();
  for (const agent of config.agents) {
    if (seen.has(agent.id)) {
      throw new Error(`Duplicate agent id: ${agent.id}`);
    }
    seen.add(agent.id);
  }
}

export function safeConfigForLog(config: BridgeConfig): unknown {
  return redactSensitive(config);
}

async function readJsonConfig(path: string): Promise<unknown> {
  return JSON.parse(await readFile(path, 'utf8')) as unknown;
}

function configFromEnv(env: NodeJS.ProcessEnv): Partial<BridgeConfig> {
  const output: Partial<BridgeConfig> = {};

  if (env.M365_ACP_HOST || env.M365_ACP_PORT) {
    output.server = {
      host: env.M365_ACP_HOST ?? '127.0.0.1',
      port: env.M365_ACP_PORT ? Number(env.M365_ACP_PORT) : 3838,
    };
  }

  if (env.M365_ACP_AUTH_MODE || env.M365_ACP_TOKEN_SOURCE) {
    output.auth = {
      mode: (env.M365_ACP_AUTH_MODE ?? 'external_token') as BridgeConfig['auth']['mode'],
      tokenSource: (env.M365_ACP_TOKEN_SOURCE ??
        'authorization_header') as BridgeConfig['auth']['tokenSource'],
    };
  }

  if (env.M365_ACP_LOG_LEVEL) {
    output.logLevel = env.M365_ACP_LOG_LEVEL as BridgeConfig['logLevel'];
  }

  const hasAgentEnv =
    env.M365_ACP_DIRECT_CONNECT_URL ||
    env.M365_ACP_CONNECTION_STRING ||
    env.M365_ACP_ENVIRONMENT_ID ||
    env.M365_ACP_SCHEMA_NAME ||
    env.M365_ACP_AGENT_IDENTIFIER;

  if (hasAgentEnv) {
    output.agents = [
      {
        id: env.M365_ACP_AGENT_ID ?? 'default',
        displayName: env.M365_ACP_AGENT_DISPLAY_NAME,
        copilotStudio: {
          tenantId: env.M365_ACP_TENANT_ID,
          environmentId: env.M365_ACP_ENVIRONMENT_ID,
          schemaName: env.M365_ACP_SCHEMA_NAME,
          agentIdentifier: env.M365_ACP_AGENT_IDENTIFIER,
          clientId: env.M365_ACP_CLIENT_ID,
          clientSecret: env.M365_ACP_CLIENT_SECRET
            ? { secretRef: 'env:M365_ACP_CLIENT_SECRET' }
            : undefined,
          connectionString: env.M365_ACP_CONNECTION_STRING
            ? { secretRef: 'env:M365_ACP_CONNECTION_STRING' }
            : undefined,
          directConnectUrl: env.M365_ACP_DIRECT_CONNECT_URL
            ? { secretRef: 'env:M365_ACP_DIRECT_CONNECT_URL' }
            : undefined,
          cloud: env.M365_ACP_CLOUD,
          customPowerPlatformCloud: env.M365_ACP_CUSTOM_POWER_PLATFORM_CLOUD,
          useExperimentalEndpoint: env.M365_ACP_USE_EXPERIMENTAL_ENDPOINT
            ? env.M365_ACP_USE_EXPERIMENTAL_ENDPOINT.toLowerCase() === 'true'
            : undefined,
        },
      },
    ];
  }

  return output;
}

function mergeConfig(base: unknown, override: unknown): unknown {
  if (!isPlainObject(base) || !isPlainObject(override)) {
    return isPlainObject(override) && Object.keys(override).length > 0 ? override : base;
  }

  const result: Record<string, unknown> = { ...base };
  for (const [key, value] of Object.entries(override)) {
    if (value === undefined) {
      continue;
    }
    result[key] = key in result ? mergeConfig(result[key], value) : value;
  }
  return result;
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
