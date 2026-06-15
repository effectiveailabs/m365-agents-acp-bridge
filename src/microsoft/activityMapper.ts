import type { SessionUpdate } from '@agentclientprotocol/sdk';
import { redactSensitive } from '../config/redact.js';
import type { MicrosoftActivity, MicrosoftCardAction, MicrosoftEntity } from './types.js';

export interface MappedActivity {
  update?: SessionUpdate;
  conversationId?: string;
  activityId?: string;
}

export function mapMicrosoftActivityToAcpUpdate(activity: MicrosoftActivity): MappedActivity {
  const microsoftMeta = buildMicrosoftMeta(activity);
  const activityType = String(activity.type ?? '').toLowerCase();
  const text = typeof activity.text === 'string' ? activity.text : '';
  const fallback = markdownFallback(activity);

  if (activityType === 'typing') {
    return {
      activityId: activity.id,
      conversationId: activity.conversation?.id,
      update: {
        sessionUpdate: 'agent_thought_chunk',
        content: {
          type: 'text',
          text: text || 'Working...',
          _meta: microsoftMeta,
        },
        messageId: activity.id,
      },
    };
  }

  if (text || fallback) {
    return {
      activityId: activity.id,
      conversationId: activity.conversation?.id,
      update: {
        sessionUpdate: 'agent_message_chunk',
        content: {
          type: 'text',
          text: [text, fallback].filter(Boolean).join('\n\n'),
          _meta: microsoftMeta,
        },
        messageId: activity.id,
      },
    };
  }

  return {
    activityId: activity.id,
    conversationId: activity.conversation?.id,
  };
}

function buildMicrosoftMeta(activity: MicrosoftActivity): { microsoft: Record<string, unknown> } {
  const meta: Record<string, unknown> = {
    activity: redactSensitive(boundPayload(activity)),
  };

  if (activity.attachments) {
    meta.attachments = redactSensitive(boundPayload(activity.attachments));
  }

  const citations = extractCitations(activity);
  if (citations.length > 0) {
    meta.citations = citations;
  }

  if (activity.suggestedActions?.actions) {
    meta.suggestedActions = redactSensitive(boundPayload(activity.suggestedActions.actions));
  }

  const auth = extractAuthPayload(activity);
  if (auth) {
    meta.auth = redactSensitive(boundPayload(auth));
  }

  return { microsoft: meta };
}

function markdownFallback(activity: MicrosoftActivity): string {
  const parts: string[] = [];

  const attachmentFallback = attachmentsMarkdown(activity.attachments);
  if (attachmentFallback) {
    parts.push(attachmentFallback);
  }

  const citations = citationsMarkdown(extractCitations(activity));
  if (citations) {
    parts.push(citations);
  }

  const suggestedActions = suggestedActionsMarkdown(activity.suggestedActions?.actions);
  if (suggestedActions) {
    parts.push(suggestedActions);
  }

  if (extractAuthPayload(activity)) {
    parts.push('Authentication is required to continue with this Microsoft agent response.');
  }

  return parts.join('\n\n');
}

function attachmentsMarkdown(attachments: unknown[] | undefined): string {
  if (!attachments || attachments.length === 0) {
    return '';
  }

  return attachments
    .map((attachment, index) => {
      const object = isRecord(attachment) ? attachment : {};
      const title =
        stringValue(object.title) ?? stringValue(object.name) ?? `Attachment ${index + 1}`;
      const contentType = stringValue(object.contentType);
      return contentType ? `- ${title} (${contentType})` : `- ${title}`;
    })
    .join('\n');
}

function suggestedActionsMarkdown(actions: MicrosoftCardAction[] | undefined): string {
  if (!actions || actions.length === 0) {
    return '';
  }

  return actions
    .map((action) => action.title ?? action.text ?? stringValue(action.value))
    .filter((value): value is string => Boolean(value))
    .map((value) => `- ${value}`)
    .join('\n');
}

function citationsMarkdown(citations: Citation[]): string {
  if (citations.length === 0) {
    return '';
  }

  return citations
    .map((citation) => {
      if (citation.url) {
        return `- [${citation.title ?? citation.url}](${citation.url})`;
      }
      return citation.title ? `- ${citation.title}` : '';
    })
    .filter(Boolean)
    .join('\n');
}

interface Citation {
  title?: string;
  url?: string;
}

function extractCitations(activity: MicrosoftActivity): Citation[] {
  const entities = activity.entities ?? [];
  const citations: Citation[] = [];

  for (const entity of entities) {
    const maybeCitation = entity as MicrosoftEntity;
    const type = maybeCitation.type?.toLowerCase();
    if (!type?.includes('citation') && !maybeCitation.citation) {
      continue;
    }

    const citationRecord = isRecord(maybeCitation.citation)
      ? maybeCitation.citation
      : maybeCitation;
    const title =
      stringValue(citationRecord.title) ??
      stringValue(citationRecord.name) ??
      stringValue(maybeCitation.title) ??
      stringValue(maybeCitation.name);
    const url = stringValue(citationRecord.url) ?? stringValue(maybeCitation.url);
    citations.push({ title, url });
  }

  return citations;
}

function extractAuthPayload(activity: MicrosoftActivity): unknown {
  const attachments = activity.attachments ?? [];
  for (const attachment of attachments) {
    if (!isRecord(attachment)) {
      continue;
    }

    const contentType = stringValue(attachment.contentType)?.toLowerCase() ?? '';
    if (
      contentType.includes('signin') ||
      contentType.includes('oauth') ||
      contentType.includes('login')
    ) {
      return attachment;
    }
  }

  if (
    activity.name?.toLowerCase().includes('signin') ||
    activity.name?.toLowerCase().includes('oauth')
  ) {
    return activity.value ?? activity.channelData ?? activity;
  }

  return undefined;
}

function boundPayload(value: unknown, maxBytes = 48_000): unknown {
  const serialized = JSON.stringify(value);
  if (!serialized || serialized.length <= maxBytes) {
    return value;
  }

  return {
    truncated: true,
    byteLength: serialized.length,
    preview: serialized.slice(0, maxBytes),
  };
}

function stringValue(value: unknown): string | undefined {
  return typeof value === 'string' && value.length > 0 ? value : undefined;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
