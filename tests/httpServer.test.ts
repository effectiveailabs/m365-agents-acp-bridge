import { afterEach, describe, expect, it } from 'vitest';
import type { StartedHttpBridgeServer } from '../src/server/httpServer.js';
import { startHttpBridgeServer } from '../src/server/httpServer.js';
import { FakeMicrosoftAgentAdapter } from '../src/microsoft/fakeAdapter.js';
import { testConfig } from './helpers.js';

describe('HTTP/SSE server', () => {
  let started: StartedHttpBridgeServer | undefined;

  afterEach(async () => {
    await started?.close();
    started = undefined;
  });

  it('serves health, JSON-RPC requests, and streamed session/update events', async () => {
    const fake = new FakeMicrosoftAgentAdapter({
      promptEvents: [
        {
          activity: {
            id: 'http-activity-1',
            type: 'message',
            text: 'HTTP bridge response',
            conversation: { id: 'http-conversation-1' },
          },
        },
      ],
    });

    started = await startHttpBridgeServer({
      config: testConfig(),
      microsoft: fake,
    });

    const health = await fetch(`${started.url}/healthz`);
    expect(await health.json()).toEqual({ ok: true });

    const stream = await fetch(`${started.url}/acp`);
    expect(stream.ok).toBe(true);
    const reader = stream.body?.getReader();
    expect(reader).toBeDefined();

    const init = await postRpc(started.url, {
      jsonrpc: '2.0',
      id: 1,
      method: 'initialize',
      params: { protocolVersion: 1 },
    });
    expect(init.result.protocolVersion).toBe(1);

    const newSession = await postRpc(started.url, {
      jsonrpc: '2.0',
      id: 2,
      method: 'session/new',
      params: { cwd: '/tmp', mcpServers: [] },
    });
    expect(newSession.result.sessionId).toBe('session-1');

    const prompt = await postRpc(
      started.url,
      {
        jsonrpc: '2.0',
        id: 3,
        method: 'session/prompt',
        params: {
          sessionId: newSession.result.sessionId,
          prompt: [{ type: 'text', text: 'Hello over HTTP' }],
        },
      },
      { authorization: 'Bearer short-lived-token' },
    );
    expect(prompt.result.stopReason).toBe('end_turn');

    const event = await readSseJson(reader!);
    expect(event).toMatchObject({
      jsonrpc: '2.0',
      method: 'session/update',
      params: {
        sessionId: 'session-1',
        update: {
          sessionUpdate: 'agent_message_chunk',
          content: {
            type: 'text',
            text: 'HTTP bridge response',
          },
        },
      },
    });

    await reader?.cancel();
  });
});

async function postRpc(
  url: string,
  body: Record<string, unknown>,
  headers: Record<string, string> = {},
): Promise<any> {
  const response = await fetch(`${url}/acp`, {
    method: 'POST',
    headers: {
      'content-type': 'application/json',
      ...headers,
    },
    body: JSON.stringify(body),
  });
  expect(response.ok).toBe(true);
  return response.json();
}

async function readSseJson(reader: ReadableStreamDefaultReader<Uint8Array>): Promise<any> {
  const decoder = new TextDecoder();
  let buffer = '';

  for (;;) {
    const { done, value } = await reader.read();
    if (done) {
      throw new Error('SSE stream closed before message event');
    }

    buffer += decoder.decode(value, { stream: true });
    const events = buffer.split('\n\n');
    buffer = events.pop() ?? '';

    for (const event of events) {
      const dataLine = event.split('\n').find((line) => line.startsWith('data: '));
      if (dataLine) {
        return JSON.parse(dataLine.slice('data: '.length));
      }
    }
  }
}
