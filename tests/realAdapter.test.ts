import { afterEach, describe, expect, it, vi } from 'vitest';
import { CopilotStudioMicrosoftAdapter } from '../src/microsoft/realAdapter.js';
import { testConfig } from './helpers.js';

describe('CopilotStudioMicrosoftAdapter with simulated Microsoft SSE', () => {
  afterEach(() => {
    vi.unstubAllGlobals();
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
