import { Activity } from '@microsoft/agents-activity';
import { CopilotStudioClient } from '@microsoft/agents-copilotstudio-client';
import type { BridgeConfig } from '../config/types.js';
import { BridgeError } from '../errors.js';
import { connectionSettingsForCopilot } from './connectionSettings.js';
import type {
  MicrosoftActivityEvent,
  MicrosoftAgentAdapter,
  SendMicrosoftPromptInput,
  StartMicrosoftSessionInput,
} from './types.js';

const DEFAULT_MICROSOFT_STREAM_TIMEOUT_MS = 60_000;

export class CopilotStudioMicrosoftAdapter implements MicrosoftAgentAdapter {
  private readonly streamTimeoutMs: number;

  constructor(
    private readonly config: BridgeConfig,
    private readonly env: NodeJS.ProcessEnv = process.env,
  ) {
    this.streamTimeoutMs = streamTimeoutMsFromEnv(env);
  }

  async startSession(input: StartMicrosoftSessionInput) {
    try {
      const client = await this.createClient(input);
      let conversationId: string | undefined;
      let activityId: string | undefined;

      for await (const activity of withMicrosoftStreamTimeout(
        client.startConversationStreaming(true),
        this.streamTimeoutMs,
      )) {
        conversationId = activity.conversation?.id ?? conversationId;
        activityId = activity.id ?? activityId;
        if (conversationId) {
          break;
        }
      }

      return { conversationId, activityId };
    } catch (error) {
      throw mapMicrosoftClientError(error);
    }
  }

  async *sendPrompt(input: SendMicrosoftPromptInput): AsyncIterable<MicrosoftActivityEvent> {
    try {
      if (input.signal?.aborted) {
        return;
      }

      const client = await this.createClient(input);
      const text = input.prompt
        .filter((block) => block.type === 'text')
        .map((block) => block.text)
        .join('\n\n');

      if (!text.trim()) {
        throw new BridgeError(
          'MS_UNSUPPORTED_ACTIVITY',
          'Only text prompts are supported in v1',
          400,
        );
      }

      const activity = Activity.fromObject({
        type: 'message',
        text,
        conversation: input.session.conversationId
          ? { id: input.session.conversationId }
          : undefined,
      });

      const stream = input.session.conversationId
        ? client.sendActivityStreaming(activity, input.session.conversationId)
        : client.sendActivityStreaming(activity);

      for await (const responseActivity of withMicrosoftStreamTimeout(
        stream,
        this.streamTimeoutMs,
      )) {
        if (input.signal?.aborted) {
          return;
        }
        yield { activity: JSON.parse(responseActivity.toJsonString()) as Record<string, unknown> };
      }
    } catch (error) {
      throw mapMicrosoftClientError(error);
    }
  }

  async cancel(): Promise<void> {
    return undefined;
  }

  private async createClient(input: StartMicrosoftSessionInput | SendMicrosoftPromptInput) {
    const token = input.auth.accessToken;
    if (!token) {
      throw new BridgeError(
        'MS_AUTH_REQUIRED',
        'A delegated Microsoft access token is required',
        401,
      );
    }

    const settings = await connectionSettingsForCopilot(input.agent.copilotStudio, this.env);
    return new CopilotStudioClient(settings, token);
  }
}

async function* withMicrosoftStreamTimeout<T>(
  stream: AsyncIterable<T>,
  timeoutMs: number,
): AsyncIterable<T> {
  const iterator = stream[Symbol.asyncIterator]();
  try {
    for (;;) {
      const next = await nextWithTimeout(iterator, timeoutMs);
      if (next.done) {
        return;
      }
      yield next.value;
    }
  } finally {
    await closeIteratorBestEffort(iterator, timeoutMs);
  }
}

async function nextWithTimeout<T>(
  iterator: AsyncIterator<T>,
  timeoutMs: number,
): Promise<IteratorResult<T>> {
  let timeout: ReturnType<typeof setTimeout> | undefined;
  const next = iterator.next();
  next.catch(() => undefined);

  try {
    return await Promise.race([
      next,
      new Promise<never>((_, reject) => {
        timeout = setTimeout(() => {
          reject(
            new BridgeError(
              'MS_STREAM_INTERRUPTED',
              `Microsoft Copilot Studio stream timed out after ${timeoutMs}ms`,
              504,
            ),
          );
        }, timeoutMs);
      }),
    ]);
  } finally {
    if (timeout) {
      clearTimeout(timeout);
    }
  }
}

async function closeIteratorBestEffort<T>(
  iterator: AsyncIterator<T>,
  timeoutMs: number,
): Promise<void> {
  if (!iterator.return) {
    return;
  }

  const closeTimeoutMs = Math.min(1000, Math.max(10, timeoutMs));
  await Promise.race([
    iterator.return(),
    new Promise((resolve) => setTimeout(resolve, closeTimeoutMs)),
  ]).catch(() => undefined);
}

function streamTimeoutMsFromEnv(env: NodeJS.ProcessEnv): number {
  const raw = env.M365_ACP_MICROSOFT_STREAM_TIMEOUT_MS;
  if (!raw) {
    return DEFAULT_MICROSOFT_STREAM_TIMEOUT_MS;
  }

  const parsed = Number(raw);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : DEFAULT_MICROSOFT_STREAM_TIMEOUT_MS;
}

function mapMicrosoftClientError(error: unknown): BridgeError {
  if (error instanceof BridgeError) {
    return error;
  }

  const message = error instanceof Error ? error.message : String(error);
  const normalized = message.toLowerCase();

  if (normalized.includes('401') || normalized.includes('unauthorized')) {
    return new BridgeError(
      'MS_AUTH_EXPIRED',
      'Microsoft Copilot Studio rejected the delegated access token',
      401,
    );
  }

  if (normalized.includes('403') || normalized.includes('forbidden')) {
    return new BridgeError(
      'MS_INVOKE_FORBIDDEN',
      'Microsoft Copilot Studio invocation was forbidden',
      403,
    );
  }

  if (normalized.includes('429') || normalized.includes('rate limit')) {
    return new BridgeError('MS_RATE_LIMITED', 'Microsoft Copilot Studio rate limit exceeded', 429);
  }

  if (
    normalized.includes('stream') ||
    normalized.includes('network') ||
    normalized.includes('fetch') ||
    normalized.includes('econnreset')
  ) {
    return new BridgeError(
      'MS_STREAM_INTERRUPTED',
      'Microsoft Copilot Studio stream interrupted',
      502,
    );
  }

  return new BridgeError('MS_INVOKE_FAILED', 'Microsoft Copilot Studio invocation failed', 502);
}
