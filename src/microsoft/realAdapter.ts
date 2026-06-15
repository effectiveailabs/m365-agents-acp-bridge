import { Activity } from '@microsoft/agents-activity';
import { ConnectionSettings, CopilotStudioClient } from '@microsoft/agents-copilotstudio-client';
import { resolveSecretRef } from '../config/secrets.js';
import type { BridgeConfig } from '../config/types.js';
import { BridgeError } from '../errors.js';
import type {
  MicrosoftActivityEvent,
  MicrosoftAgentAdapter,
  SendMicrosoftPromptInput,
  StartMicrosoftSessionInput,
} from './types.js';

export class CopilotStudioMicrosoftAdapter implements MicrosoftAgentAdapter {
  constructor(
    private readonly config: BridgeConfig,
    private readonly env: NodeJS.ProcessEnv = process.env,
  ) {}

  async startSession(input: StartMicrosoftSessionInput) {
    try {
      const client = await this.createClient(input);
      let conversationId: string | undefined;
      let activityId: string | undefined;

      for await (const activity of client.startConversationStreaming(true)) {
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

      for await (const responseActivity of stream) {
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

    const settings = await this.connectionSettings(input);
    return new CopilotStudioClient(settings, token);
  }

  private async connectionSettings(input: StartMicrosoftSessionInput | SendMicrosoftPromptInput) {
    const copilot = input.agent.copilotStudio;
    const directConnectUrl = copilot.directConnectUrl
      ? await resolveSecretRef(copilot.directConnectUrl, this.env)
      : undefined;

    const connectionString = copilot.connectionString
      ? await resolveSecretRef(copilot.connectionString, this.env)
      : undefined;

    return new ConnectionSettings({
      tenantId: copilot.tenantId,
      appClientId: copilot.clientId,
      environmentId: copilot.environmentId,
      schemaName: copilot.schemaName,
      agentIdentifier: copilot.agentIdentifier,
      directConnectUrl: directConnectUrl ?? connectionString,
      cloud: copilot.cloud as ConstructorParameters<typeof ConnectionSettings>[0]['cloud'],
      customPowerPlatformCloud: copilot.customPowerPlatformCloud,
      useExperimentalEndpoint: copilot.useExperimentalEndpoint,
    });
  }
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
