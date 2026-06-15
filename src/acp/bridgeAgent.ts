import type {
  Agent,
  AuthenticateRequest,
  AuthenticateResponse,
  CancelNotification,
  InitializeRequest,
  InitializeResponse,
  NewSessionRequest,
  NewSessionResponse,
  PromptRequest,
  PromptResponse,
} from '@agentclientprotocol/sdk';
import type { BridgeConfig, AgentConfig } from '../config/types.js';
import { BridgeError } from '../errors.js';
import { consoleLogger, type Logger } from '../logging/logger.js';
import { mapMicrosoftActivityToAcpUpdate } from '../microsoft/activityMapper.js';
import type { MicrosoftAgentAdapter, RequestAuthContext } from '../microsoft/types.js';
import { MemorySessionStore, type SessionStore } from '../session/memorySessionStore.js';
import type { AcpUpdateSink } from './updateSink.js';

export interface BridgeAgentOptions {
  config: BridgeConfig;
  microsoft: MicrosoftAgentAdapter;
  updates: AcpUpdateSink;
  sessions?: SessionStore;
  logger?: Logger;
  defaultAgentId?: string;
}

export class BridgeAgent implements Agent {
  private readonly config: BridgeConfig;
  private readonly microsoft: MicrosoftAgentAdapter;
  private readonly updates: AcpUpdateSink;
  private readonly sessions: SessionStore;
  private readonly logger: Logger;
  private readonly defaultAgentId?: string;
  private auth: RequestAuthContext = {};
  private readonly activeTurns = new Map<string, AbortController>();

  constructor(options: BridgeAgentOptions) {
    this.config = options.config;
    this.microsoft = options.microsoft;
    this.updates = options.updates;
    this.sessions = options.sessions ?? new MemorySessionStore();
    this.logger = options.logger ?? consoleLogger;
    this.defaultAgentId = options.defaultAgentId;
  }

  async initialize(params: InitializeRequest): Promise<InitializeResponse> {
    return {
      protocolVersion: params.protocolVersion,
      agentInfo: {
        name: 'm365-agents-acp-bridge',
        title: 'M365 Agents ACP Bridge',
        version: '0.1.0',
      },
      agentCapabilities: {
        promptCapabilities: {},
        sessionCapabilities: {},
        auth: {},
        _meta: {
          microsoft: {
            provider: 'copilot_studio',
          },
        },
      },
      authMethods: authMethodsForMode(this.config.auth.mode),
    };
  }

  async authenticate(params: AuthenticateRequest): Promise<AuthenticateResponse> {
    const token = tokenFromAuthenticateMeta(params);
    this.auth = {
      accessToken: token,
      authMethodId: params.methodId,
      meta: params._meta ?? undefined,
    };
    return {};
  }

  async newSession(params: NewSessionRequest): Promise<NewSessionResponse> {
    const agent = this.selectAgent(params._meta);
    const microsoftSession = await this.microsoft.startSession({
      agent,
      auth: this.auth,
    });
    const session = await this.sessions.create({
      agentId: agent.id,
      microsoft: microsoftSession,
    });

    this.logger.info('created ACP session', {
      agentId: agent.id,
      sessionId: session.sessionId,
      microsoftConversationId: microsoftSession.conversationId,
    });

    return {
      sessionId: session.sessionId,
      _meta: {
        microsoft: {
          conversationId: microsoftSession.conversationId,
          activityId: microsoftSession.activityId,
        },
      },
    };
  }

  async prompt(params: PromptRequest): Promise<PromptResponse> {
    const session = await this.sessions.get(params.sessionId);
    if (!session) {
      throw new BridgeError('MS_SESSION_NOT_FOUND', `Session not found: ${params.sessionId}`, 404);
    }

    const agent = this.agentById(session.agentId);
    const abortController = new AbortController();
    this.activeTurns.set(params.sessionId, abortController);

    try {
      for await (const event of this.microsoft.sendPrompt({
        agent,
        session: session.microsoft,
        prompt: params.prompt,
        auth: this.auth,
        signal: abortController.signal,
      })) {
        if (abortController.signal.aborted) {
          break;
        }

        const mapped = mapMicrosoftActivityToAcpUpdate(event.activity);
        if (mapped.conversationId || mapped.activityId) {
          await this.sessions.updateMicrosoft(params.sessionId, {
            conversationId: mapped.conversationId,
            activityId: mapped.activityId,
          });
        }

        if (mapped.update && !abortController.signal.aborted) {
          await this.updates.sessionUpdate({
            sessionId: params.sessionId,
            update: mapped.update,
          });
        }
      }

      return {
        stopReason: abortController.signal.aborted ? 'cancelled' : 'end_turn',
      };
    } finally {
      this.activeTurns.delete(params.sessionId);
    }
  }

  async cancel(params: CancelNotification): Promise<void> {
    const active = this.activeTurns.get(params.sessionId);
    active?.abort();

    const session = await this.sessions.get(params.sessionId);
    if (!session) {
      return;
    }

    const agent = this.agentById(session.agentId);
    await this.microsoft.cancel({
      agent,
      session: session.microsoft,
      reason: 'session/cancel',
    });
  }

  async loadSession(): Promise<never> {
    throw new BridgeError('MS_SESSION_NOT_FOUND', 'session/load is not supported in v1', 404);
  }

  async unstable_forkSession(): Promise<never> {
    throw new BridgeError('MS_SESSION_NOT_FOUND', 'session/fork is not supported in v1', 404);
  }

  async listSessions(): Promise<never> {
    throw new BridgeError('MS_SESSION_NOT_FOUND', 'session/list is not supported in v1', 404);
  }

  async deleteSession(): Promise<never> {
    throw new BridgeError('MS_SESSION_NOT_FOUND', 'session/delete is not supported in v1', 404);
  }

  async resumeSession(): Promise<never> {
    throw new BridgeError('MS_SESSION_NOT_FOUND', 'session/resume is not supported in v1', 404);
  }

  async closeSession(): Promise<never> {
    throw new BridgeError('MS_SESSION_NOT_FOUND', 'session/close is not supported in v1', 404);
  }

  async setSessionMode(): Promise<never> {
    throw new BridgeError(
      'MS_UNSUPPORTED_ACTIVITY',
      'session/set_mode is not supported in v1',
      400,
    );
  }

  async setSessionConfigOption(): Promise<never> {
    throw new BridgeError(
      'MS_UNSUPPORTED_ACTIVITY',
      'session/set_config_option is not supported in v1',
      400,
    );
  }

  async unstable_listProviders(): Promise<never> {
    throw new BridgeError('MS_UNSUPPORTED_ACTIVITY', 'providers/list is not supported in v1', 400);
  }

  async unstable_setProvider(): Promise<never> {
    throw new BridgeError('MS_UNSUPPORTED_ACTIVITY', 'providers/set is not supported in v1', 400);
  }

  async unstable_disableProvider(): Promise<never> {
    throw new BridgeError(
      'MS_UNSUPPORTED_ACTIVITY',
      'providers/disable is not supported in v1',
      400,
    );
  }

  async logout(): Promise<Record<string, never>> {
    this.auth = {};
    return {};
  }

  async unstable_startNes(): Promise<never> {
    throw new BridgeError('MS_UNSUPPORTED_ACTIVITY', 'nes/start is not supported in v1', 400);
  }

  async unstable_suggestNes(): Promise<never> {
    throw new BridgeError('MS_UNSUPPORTED_ACTIVITY', 'nes/suggest is not supported in v1', 400);
  }

  async unstable_closeNes(): Promise<never> {
    throw new BridgeError('MS_UNSUPPORTED_ACTIVITY', 'nes/close is not supported in v1', 400);
  }

  async unstable_didOpenDocument(): Promise<void> {}
  async unstable_didChangeDocument(): Promise<void> {}
  async unstable_didCloseDocument(): Promise<void> {}
  async unstable_didSaveDocument(): Promise<void> {}
  async unstable_didFocusDocument(): Promise<void> {}
  async unstable_acceptNes(): Promise<void> {}
  async unstable_rejectNes(): Promise<void> {}

  async extMethod(): Promise<Record<string, unknown>> {
    throw new BridgeError(
      'MS_UNSUPPORTED_ACTIVITY',
      'ACP extension methods are not supported in v1',
      400,
    );
  }

  async extNotification(): Promise<void> {}

  private selectAgent(meta: Record<string, unknown> | null | undefined): AgentConfig {
    const metaAgentId = isRecord(meta) ? stringValue(meta.agentId) : undefined;
    const agentId = metaAgentId ?? this.defaultAgentId ?? this.config.agents[0]?.id;
    if (!agentId) {
      throw new BridgeError('MS_AGENT_NOT_CONFIGURED', 'No Microsoft agent is configured', 500);
    }
    return this.agentById(agentId);
  }

  private agentById(agentId: string): AgentConfig {
    const agent = this.config.agents.find((candidate) => candidate.id === agentId);
    if (!agent) {
      throw new BridgeError(
        'MS_AGENT_NOT_FOUND',
        `Microsoft agent is not configured: ${agentId}`,
        404,
      );
    }
    return agent;
  }
}

function authMethodsForMode(mode: BridgeConfig['auth']['mode']) {
  if (mode === 'interactive') {
    return [
      {
        id: 'interactive',
        name: 'Microsoft interactive login',
        description: 'Use the bridge host interactive Microsoft login flow.',
      },
    ];
  }

  return [
    {
      id: 'external_token',
      name: 'Delegated Microsoft access token',
      description: 'Provide a short-lived delegated Microsoft access token.',
      _meta: {
        microsoft: {
          permission: 'CopilotStudio.Copilots.Invoke',
        },
      },
    },
  ];
}

function tokenFromAuthenticateMeta(params: AuthenticateRequest): string | undefined {
  const meta = params._meta;
  if (!isRecord(meta)) {
    return undefined;
  }

  const token = meta.accessToken ?? meta.access_token ?? meta.microsoftAccessToken;
  return typeof token === 'string' ? token : undefined;
}

function stringValue(value: unknown): string | undefined {
  return typeof value === 'string' && value.length > 0 ? value : undefined;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
