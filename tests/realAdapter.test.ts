import { afterEach, describe, expect, it, vi } from 'vitest';
import { CopilotStudioClient } from '@microsoft/agents-copilotstudio-client';
import { CopilotStudioMicrosoftAdapter } from '../src/microsoft/realAdapter.js';
import type { SendMicrosoftPromptInput } from '../src/microsoft/types.js';
import { testConfig } from './helpers.js';

describe('CopilotStudioMicrosoftAdapter with simulated Microsoft SSE', () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it('consumes Microsoft-like SSE from the real SDK client without tenant credentials', async () => {
    const fetchMock = vi.fn(async () =>
      sseResponse({
        id: 'start-activity',
        type: 'message',
        text: 'Started',
        conversation: { id: 'conversation-1' },
      }),
    );
    vi.stubGlobal('fetch', fetchMock);

    const config = testConfig();
    const adapter = new CopilotStudioMicrosoftAdapter(config, {
      COPILOT_STUDIO_DIRECT_CONNECT_URL: 'https://copilot.example/direct',
    });

    const agent = config.agents[0]!;
    const session = await adapter.startSession({
      agent,
      auth: { accessToken: 'short-lived-token' },
    });

    expect(session.conversationId).toBe('conversation-1');
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it('streams prompt response activities through the real SDK client', async () => {
    const fetchMock = vi.fn(async () =>
      sseResponse({
        id: 'prompt-activity',
        type: 'message',
        text: 'Prompt response',
        conversation: { id: 'conversation-1' },
      }),
    );
    vi.stubGlobal('fetch', fetchMock);

    const config = testConfig();
    const adapter = new CopilotStudioMicrosoftAdapter(config, {
      COPILOT_STUDIO_DIRECT_CONNECT_URL: 'https://copilot.example/direct',
    });

    const events = [];
    for await (const event of adapter.sendPrompt({
      agent: config.agents[0]!,
      auth: { accessToken: 'short-lived-token' },
      session: { conversationId: 'conversation-1' },
      prompt: [{ type: 'text', text: 'Hello' }],
    })) {
      events.push(event);
    }

    expect(events).toHaveLength(1);
    expect(events[0]?.activity).toMatchObject({
      id: 'prompt-activity',
      type: 'message',
      text: 'Prompt response',
      conversation: { id: 'conversation-1' },
    });
  });

  it('does not crash on malformed Microsoft activity events', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async () =>
        rawSseResponse('event: activity\ndata: {bad-json}\n\nevent: end\ndata: {}\n\n'),
      ),
    );

    const config = testConfig();
    const adapter = new CopilotStudioMicrosoftAdapter(config, {
      COPILOT_STUDIO_DIRECT_CONNECT_URL: 'https://copilot.example/direct',
    });

    const session = await adapter.startSession({
      agent: config.agents[0]!,
      auth: { accessToken: 'short-lived-token' },
    });

    expect(session).toEqual({
      activityId: undefined,
      conversationId: undefined,
    });
  });

  it('requires a delegated Microsoft access token', async () => {
    const config = testConfig();
    const adapter = new CopilotStudioMicrosoftAdapter(config, {
      COPILOT_STUDIO_DIRECT_CONNECT_URL: 'https://copilot.example/direct',
    });

    await expect(
      adapter.startSession({
        agent: config.agents[0]!,
        auth: {},
      }),
    ).rejects.toMatchObject({
      code: 'MS_AUTH_REQUIRED',
      status: 401,
    });
  });

  it('rejects unsupported non-text prompt content before invoking Microsoft', async () => {
    const fetchMock = vi.fn();
    vi.stubGlobal('fetch', fetchMock);

    const config = testConfig();
    const adapter = new CopilotStudioMicrosoftAdapter(config, {
      COPILOT_STUDIO_DIRECT_CONNECT_URL: 'https://copilot.example/direct',
    });

    await expect(
      collectPrompt(adapter, {
        agent: config.agents[0]!,
        auth: { accessToken: 'short-lived-token' },
        session: { conversationId: 'conversation-1' },
        prompt: [{ type: 'image', data: 'abc', mimeType: 'image/png' }],
      }),
    ).rejects.toMatchObject({
      code: 'MS_UNSUPPORTED_ACTIVITY',
      status: 400,
    });
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it.each([
    ['401 Unauthorized', 'MS_AUTH_EXPIRED', 401],
    ['403 Forbidden', 'MS_INVOKE_FORBIDDEN', 403],
    ['429 Too Many Requests rate limit', 'MS_RATE_LIMITED', 429],
    ['network stream interrupted', 'MS_STREAM_INTERRUPTED', 502],
  ])('maps Microsoft client failure %s to %s', async (message, code, status) => {
    vi.spyOn(CopilotStudioClient.prototype, 'startConversationStreaming').mockImplementation(
      async function* () {
        if (message === '__never__') {
          yield undefined as never;
        }
        throw new Error(message);
      },
    );

    const config = testConfig();
    const adapter = new CopilotStudioMicrosoftAdapter(config, {
      COPILOT_STUDIO_DIRECT_CONNECT_URL: 'https://copilot.example/direct',
    });

    await expect(
      adapter.startSession({
        agent: config.agents[0]!,
        auth: { accessToken: 'short-lived-token' },
      }),
    ).rejects.toMatchObject({
      code,
      status,
    });
  });

  it('times out when the Microsoft SDK stream never yields', async () => {
    vi.spyOn(CopilotStudioClient.prototype, 'startConversationStreaming').mockImplementation(
      async function* () {
        if (Date.now() === 0) {
          yield undefined as never;
        }
        await new Promise(() => undefined);
      },
    );

    const config = testConfig();
    const adapter = new CopilotStudioMicrosoftAdapter(config, {
      COPILOT_STUDIO_DIRECT_CONNECT_URL: 'https://copilot.example/direct',
      M365_ACP_MICROSOFT_STREAM_TIMEOUT_MS: '5',
    });

    await expect(
      adapter.startSession({
        agent: config.agents[0]!,
        auth: { accessToken: 'short-lived-token' },
      }),
    ).rejects.toMatchObject({
      code: 'MS_STREAM_INTERRUPTED',
      status: 504,
    });
  });
});

function sseResponse(activity: Record<string, unknown>): Response {
  const encoder = new TextEncoder();
  const body = new ReadableStream<Uint8Array>({
    start(controller) {
      controller.enqueue(
        encoder.encode(
          `event: activity\ndata: ${JSON.stringify(activity)}\n\nevent: end\ndata: {}\n\n`,
        ),
      );
      controller.close();
    },
  });

  return new Response(body, {
    status: 200,
    headers: {
      'content-type': 'text/event-stream',
      'x-ms-conversationid': 'conversation-1',
    },
  });
}

function rawSseResponse(payload: string): Response {
  const encoder = new TextEncoder();
  const body = new ReadableStream<Uint8Array>({
    start(controller) {
      controller.enqueue(encoder.encode(payload));
      controller.close();
    },
  });

  return new Response(body, {
    status: 200,
    headers: {
      'content-type': 'text/event-stream',
      'x-ms-conversationid': 'conversation-1',
    },
  });
}

async function collectPrompt(
  adapter: CopilotStudioMicrosoftAdapter,
  input: SendMicrosoftPromptInput,
) {
  const events = [];
  for await (const event of adapter.sendPrompt(input)) {
    events.push(event);
  }
  return events;
}
