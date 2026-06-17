#!/usr/bin/env node
import { loadBridgeConfig, safeConfigForLog } from './config/load.js';
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
  m365-acp stdio --config ./m365-acp.config.json
  m365-acp init

Commands:
  serve   Start the HTTP/SSE ACP bridge
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

main().catch((error: unknown) => {
  const message = error instanceof Error ? error.message : String(error);
  consoleLogger.error(message);
  process.exitCode = 1;
});
