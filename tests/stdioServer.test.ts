import {
  ClientSideConnection,
  ndJsonStream,
  type Client,
  type SessionNotification,
} from '@agentclientprotocol/sdk';
import { describe, expect, it } from 'vitest';
import { createAcpAgentConnection } from '../src/server/stdioServer.js';
import { FakeMicrosoftAgentAdapter } from '../src/microsoft/fakeAdapter.js';
import type { Logger } from '../src/logging/logger.js';
import { testConfig } from './helpers.js';

describe('stdio ACP server', () => {
  it('runs ACP lifecycle over SDK ndjson streams', async () => {
    const clientToAgent = new TransformStream();
    const agentToClient = new TransformStream();
    const updates: SessionNotification[] = [];
    const fake = new FakeMicrosoftAgentAdapter({
      promptEvents: [
        {
          activity: {
            id: 'stdio-activity-1',
            type: 'message',
            text: 'stdio bridge response',
            conversation: { id: 'stdio-conversation-1' },
          },
        },
      ],
    });

    const client = new ClientSideConnection(
      () => testClient(updates),
      ndJsonStream(clientToAgent.writable, agentToClient.readable),
    );
    createAcpAgentConnection({
      config: testConfig(),
      microsoft: fake,
      logger: silentLogger,
      stream: ndJsonStream(agentToClient.writable, clientToAgent.readable),
    });

    const init = await client.initialize({ protocolVersion: 1 });
    expect(init.authMethods?.[0]?.id).toBe('external_token');

    await client.authenticate({
      methodId: 'external_token',
      _meta: { accessToken: 'stdio-token' },
    });

    const session = await client.newSession({ cwd: '/tmp', mcpServers: [] });
    const response = await client.prompt({
      sessionId: session.sessionId,
      prompt: [{ type: 'text', text: 'Hello over stdio' }],
    });

    expect(response.stopReason).toBe('end_turn');
    expect(fake.startSessionCalls[0]?.auth.accessToken).toBe('stdio-token');
    expect(fake.sendPromptCalls[0]?.auth.accessToken).toBe('stdio-token');
    expect(updates).toHaveLength(1);
    expect(updates[0]).toMatchObject({
      sessionId: session.sessionId,
      update: {
        sessionUpdate: 'agent_message_chunk',
        content: {
          type: 'text',
          text: 'stdio bridge response',
        },
      },
    });
  });
});

function testClient(updates: SessionNotification[]): Client {
  return {
    async requestPermission() {
      throw new Error('bridge should not request client permissions in v1');
    },
    async sessionUpdate(update) {
      updates.push(update);
    },
    async readTextFile() {
      throw new Error('bridge should not read client files in v1');
    },
    async writeTextFile() {
      throw new Error('bridge should not write client files in v1');
    },
    async createTerminal() {
      throw new Error('bridge should not create client terminals in v1');
    },
  };
}

const silentLogger: Logger = {
  debug() {},
  info() {},
  warn() {},
  error() {},
};
