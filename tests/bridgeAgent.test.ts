import { describe, expect, it } from 'vitest';
import { BridgeAgent } from '../src/acp/bridgeAgent.js';
import { RecordingUpdateSink } from '../src/acp/updateSink.js';
import { FakeMicrosoftAgentAdapter } from '../src/microsoft/fakeAdapter.js';
import { MemorySessionStore } from '../src/session/memorySessionStore.js';
import { testConfig, wait } from './helpers.js';

describe('BridgeAgent', () => {
  it('runs initialize, authenticate, session/new, and session/prompt with fake Microsoft adapter', async () => {
    const fake = new FakeMicrosoftAgentAdapter({
      promptEvents: [
        {
          activity: {
            id: 'activity-1',
            type: 'message',
            text: 'Bridge response',
            conversation: { id: 'conversation-1' },
          },
        },
      ],
    });
    const updates = new RecordingUpdateSink();
    const agent = new BridgeAgent({
      config: testConfig(),
      microsoft: fake,
      updates,
      sessions: new MemorySessionStore(),
    });

    const init = await agent.initialize({ protocolVersion: 1 });
    expect(init.agentCapabilities?.promptCapabilities).toEqual({});
    expect(init.authMethods?.[0]?.id).toBe('external_token');

    await agent.authenticate({
      methodId: 'external_token',
      _meta: { accessToken: 'short-lived-token' },
    });

    const session = await agent.newSession({ cwd: '/tmp', mcpServers: [] });
    expect(session.sessionId).toBe('session-1');
    expect(fake.startSessionCalls).toHaveLength(1);

    const response = await agent.prompt({
      sessionId: session.sessionId,
      prompt: [{ type: 'text', text: 'Hello' }],
    });

    expect(response.stopReason).toBe('end_turn');
    expect(fake.sendPromptCalls).toHaveLength(1);
    expect(fake.sendPromptCalls[0]?.auth.accessToken).toBe('short-lived-token');
    expect(updates.updates).toHaveLength(1);
    expect(updates.updates[0]).toMatchObject({
      sessionId: session.sessionId,
      update: {
        sessionUpdate: 'agent_message_chunk',
        content: { type: 'text', text: 'Bridge response' },
      },
    });
  });

  it('uses per-request auth metadata instead of leaking the previous authenticate token', async () => {
    const fake = new FakeMicrosoftAgentAdapter();
    const agent = new BridgeAgent({
      config: testConfig(),
      microsoft: fake,
      updates: new RecordingUpdateSink(),
      sessions: new MemorySessionStore(),
    });

    await agent.authenticate({
      methodId: 'external_token',
      _meta: { accessToken: 'old-token' },
    });

    const session = await agent.newSession({
      cwd: '/tmp',
      mcpServers: [],
      _meta: { accessToken: 'new-session-token' },
    });
    await agent.prompt({
      sessionId: session.sessionId,
      prompt: [{ type: 'text', text: 'Hello' }],
      _meta: { accessToken: 'new-prompt-token' },
    });

    expect(fake.startSessionCalls[0]?.auth.accessToken).toBe('new-session-token');
    expect(fake.sendPromptCalls[0]?.auth.accessToken).toBe('new-prompt-token');
  });

  it('treats session/cancel as best-effort notification and suppresses late fake events', async () => {
    const fake = new FakeMicrosoftAgentAdapter({
      promptEvents: [
        {
          delayMs: 75,
          activity: {
            id: 'late-activity',
            type: 'message',
            text: 'This should not be forwarded',
            conversation: { id: 'conversation-1' },
          },
        },
      ],
    });
    const updates = new RecordingUpdateSink();
    const agent = new BridgeAgent({
      config: testConfig(),
      microsoft: fake,
      updates,
      sessions: new MemorySessionStore(),
    });

    const session = await agent.newSession({ cwd: '/tmp', mcpServers: [] });
    const prompt = agent.prompt({
      sessionId: session.sessionId,
      prompt: [{ type: 'text', text: 'Cancel me' }],
    });

    await wait(10);
    await agent.cancel({ sessionId: session.sessionId });

    await expect(prompt).resolves.toEqual({ stopReason: 'cancelled' });
    expect(fake.cancelCalls).toHaveLength(1);
    expect(updates.updates).toHaveLength(0);
  });
});
