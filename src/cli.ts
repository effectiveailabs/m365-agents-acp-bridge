#!/usr/bin/env node
import { accessTokenFromEnv } from './auth/token.js';
import { loadBridgeConfig, safeConfigForLog } from './config/load.js';
import { formatDoctorResult, runDoctor } from './diagnostics/doctor.js';
import { formatProbeResult, runProbe } from './diagnostics/probe.js';
import { consoleLogger, stderrLogger } from './logging/logger.js';
import { CopilotStudioMicrosoftAdapter } from './microsoft/realAdapter.js';
import { startHttpBridgeServer } from './server/httpServer.js';
import { runStdioBridgeServer } from './server/stdioServer.js';

async function main(): Promise<void> {
  const [command, ...args] = process.argv.slice(2);

  if (command === '--help' || command === '-h' || !command) {
    printHelp();
    return;
  }

  if (command === 'init') {
    printInitTemplate();
    return;
  }

  if (command === 'doctor') {
    const configFile = valueAfter(args, '--config');
    const tokenEnv = valueAfter(args, '--token-env');
    const config = await loadBridgeConfig({ configFile });
    const result = await runDoctor({
      config,
      accessToken: accessTokenFromEnv(process.env, tokenEnv),
    });
    console.log(formatDoctorResult(result));
    if (!result.ok) {
      process.exitCode = 1;
    }
    return;
  }

  if (command === 'probe') {
    const configFile = valueAfter(args, '--config');
    const tokenEnv = valueAfter(args, '--token-env');
    const transcriptFile = valueAfter(args, '--transcript');
    const timeoutMs = numberAfter(args, '--timeout-ms');
    const prompt = valueAfter(args, '--prompt');
    const config = configFile ? await loadBridgeConfig({ configFile }) : undefined;
    const url = valueAfter(args, '--url') ?? urlFromConfigOrEnv(config, process.env);
    const result = await runProbe({
      url,
      accessToken: accessTokenFromEnv(process.env, tokenEnv),
      prompt,
      timeoutMs,
      transcriptFile,
    });
    console.log(formatProbeResult(result));
    if (!result.ok) {
      process.exitCode = 1;
    }
    return;
  }

  if (command === 'stdio') {
    const configFile = valueAfter(args, '--config');
    const config = await loadBridgeConfig({ configFile });
    stderrLogger.debug('loaded config', {
      config: safeConfigForLog(config) as Record<string, unknown>,
    });
    await runStdioBridgeServer({
      config,
      microsoft: new CopilotStudioMicrosoftAdapter(config),
      logger: stderrLogger,
    });
    return;
  }

  if (command !== 'serve') {
    throw new Error(`Unknown command: ${command}`);
  }

  const configFile = valueAfter(args, '--config');
  const config = await loadBridgeConfig({ configFile });
  consoleLogger.debug('loaded config', {
    config: safeConfigForLog(config) as Record<string, unknown>,
  });

  const server = await startHttpBridgeServer({
    config,
    microsoft: new CopilotStudioMicrosoftAdapter(config),
    logger: consoleLogger,
  });

  process.on('SIGINT', () => {
    void server.close().then(() => process.exit(0));
  });
  process.on('SIGTERM', () => {
    void server.close().then(() => process.exit(0));
  });
}

function printHelp(): void {
  console.log(`M365 Agents ACP Bridge

Usage:
  m365-acp serve --config ./m365-acp.config.json
  m365-acp doctor --config ./m365-acp.config.json
  m365-acp probe --url http://127.0.0.1:3838/acp --prompt "Hello"
  m365-acp stdio --config ./m365-acp.config.json
  m365-acp init

Commands:
  serve   Start the HTTP/SSE ACP bridge
  doctor  Validate local config and delegated Microsoft token shape
  probe   Exercise initialize/session/new/session/prompt over HTTP/SSE
  stdio   Start the ACP bridge over stdio NDJSON
  init    Print an example config
`);
}

function printInitTemplate(): void {
  console.log(
    JSON.stringify(
      {
        server: {
          host: '127.0.0.1',
          port: 3838,
        },
        auth: {
          mode: 'external_token',
          tokenSource: 'authorization_header',
        },
        agents: [
          {
            id: 'default',
            displayName: 'Copilot Studio Agent',
            copilotStudio: {
              directConnectUrl: {
                secretRef: 'env:COPILOT_STUDIO_DIRECT_CONNECT_URL',
              },
            },
          },
        ],
      },
      null,
      2,
    ),
  );
}

function valueAfter(args: string[], key: string): string | undefined {
  const index = args.indexOf(key);
  if (index < 0) {
    return undefined;
  }
  return args[index + 1];
}

function numberAfter(args: string[], key: string): number | undefined {
  const value = valueAfter(args, key);
  if (!value) {
    return undefined;
  }

  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed <= 0) {
    throw new Error(`${key} must be a positive number`);
  }
  return parsed;
}

function urlFromConfigOrEnv(
  config: { server: { host: string; port: number } } | undefined,
  env: NodeJS.ProcessEnv,
): string {
  const host = config?.server.host ?? env.M365_ACP_HOST ?? '127.0.0.1';
  const port = config?.server.port ?? (env.M365_ACP_PORT ? Number(env.M365_ACP_PORT) : 3838);
  return `http://${host}:${port}/acp`;
}

main().catch((error: unknown) => {
  const message = error instanceof Error ? error.message : String(error);
  consoleLogger.error(message);
  process.exitCode = 1;
});
