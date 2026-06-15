import { describe, expect, it } from 'vitest';
import { mapMicrosoftActivityToAcpUpdate } from '../src/microsoft/activityMapper.js';

describe('activity mapper', () => {
  it('maps text messages to ACP agent message chunks', () => {
    const mapped = mapMicrosoftActivityToAcpUpdate({
      id: 'activity-1',
      type: 'message',
      text: 'Hello from Microsoft',
      conversation: { id: 'conversation-1' },
    });

    expect(mapped.conversationId).toBe('conversation-1');
    expect(mapped.update?.sessionUpdate).toBe('agent_message_chunk');
    expect(mapped.update).toMatchObject({
      content: {
        type: 'text',
        text: 'Hello from Microsoft',
      },
      messageId: 'activity-1',
    });
  });

  it('preserves rich Microsoft payloads under _meta.microsoft and emits markdown fallback', () => {
    const mapped = mapMicrosoftActivityToAcpUpdate({
      id: 'activity-rich',
      type: 'message',
      text: 'Here are the options.',
      conversation: { id: 'conversation-rich' },
      attachments: [
        {
          contentType: 'application/vnd.microsoft.card.adaptive',
          title: 'Account summary',
          content: { type: 'AdaptiveCard', body: [{ type: 'TextBlock', text: 'Summary' }] },
        },
        {
          contentType: 'application/vnd.microsoft.card.signin',
          title: 'Sign in',
          content: { buttons: [{ type: 'signin', value: 'https://login.example' }] },
        },
      ],
      entities: [
        {
          type: 'citation',
          title: 'Source Doc',
          url: 'https://example.com/source',
        },
      ],
      suggestedActions: {
        actions: [{ title: 'Approve' }, { title: 'Reject' }],
      },
    });

    expect(mapped.update?.sessionUpdate).toBe('agent_message_chunk');
    if (mapped.update?.sessionUpdate !== 'agent_message_chunk') {
      throw new Error('expected agent_message_chunk');
    }
    const content = mapped.update.content;
    expect(content?.type).toBe('text');
    if (content.type !== 'text') {
      throw new Error('expected text content');
    }
    expect(content?.text).toContain('Here are the options.');
    expect(content?.text).toContain('Account summary');
    expect(content?.text).toContain('[Source Doc](https://example.com/source)');
    expect(content?.text).toContain('Approve');
    expect(content?._meta?.microsoft).toMatchObject({
      attachments: expect.any(Array),
      citations: expect.any(Array),
      suggestedActions: expect.any(Array),
      auth: expect.any(Object),
    });
  });

  it('maps typing activities to thought chunks', () => {
    const mapped = mapMicrosoftActivityToAcpUpdate({
      id: 'typing-1',
      type: 'typing',
      text: 'Looking that up',
    });

    expect(mapped.update?.sessionUpdate).toBe('agent_thought_chunk');
    expect(mapped.update).toMatchObject({
      content: {
        type: 'text',
        text: 'Looking that up',
      },
    });
  });
});
