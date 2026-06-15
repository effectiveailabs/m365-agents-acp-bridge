import { Readable, Writable } from 'node:stream';
import { AgentSideConnection, ndJsonStream, type Stream } from '@agentclientprotocol/sdk';
import { BridgeAgent } from '../acp/bridgeAgent.js';
import type { BridgeConfig } from '../config/types.js';
import { stderrLogger, type Logger } from '../logging/logger.js';
import type { MicrosoftAgentAdapter } from '../microsoft/types.js';

export interface StdioBridgeServerOptions {
  config: BridgeConfig;
  microsoft: MicrosoftAgentAdapter;
  logger?: Logger;
  stream?: Stream;
}

export function createAcpAgentConnection(options: StdioBridgeServerOptions): AgentSideConnection {
  const stream = options.stream ?? nodeStdioStream();

  return new AgentSideConnection(
    (connection) =>
      new BridgeAgent({
        config: options.config,
        microsoft: options.microsoft,
        updates: {
          sessionUpdate: (update) => connection.sessionUpdate(update),
        },
        logger: options.logger ?? stderrLogger,
      }),
    stream,
  );
}

export async function runStdioBridgeServer(options: StdioBridgeServerOptions): Promise<void> {
  const connection = createAcpAgentConnection(options);
  await connection.closed;
}

function nodeStdioStream(): Stream {
  return ndJsonStream(
    Writable.toWeb(process.stdout) as WritableStream<Uint8Array>,
    Readable.toWeb(process.stdin) as ReadableStream<Uint8Array>,
  );
}
