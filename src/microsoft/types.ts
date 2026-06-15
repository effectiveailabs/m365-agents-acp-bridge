import type { ContentBlock } from '@agentclientprotocol/sdk';
import type { AgentConfig, BridgeConfig } from '../config/types.js';

export interface RequestAuthContext {
  accessToken?: string;
  authMethodId?: string;
  meta?: Record<string, unknown>;
}

export interface MicrosoftSession {
  conversationId?: string;
  activityId?: string;
  metadata?: Record<string, unknown>;
}

export interface StartMicrosoftSessionInput {
  agent: AgentConfig;
  auth: RequestAuthContext;
  signal?: AbortSignal;
}

export interface SendMicrosoftPromptInput {
  agent: AgentConfig;
  session: MicrosoftSession;
  prompt: ContentBlock[];
  auth: RequestAuthContext;
  signal?: AbortSignal;
}

export interface CancelMicrosoftTurnInput {
  agent: AgentConfig;
  session: MicrosoftSession;
  reason?: string;
}

export interface MicrosoftActivityEvent {
  activity: MicrosoftActivity;
}

export interface MicrosoftAgentAdapter {
  startSession(input: StartMicrosoftSessionInput): Promise<MicrosoftSession>;
  sendPrompt(input: SendMicrosoftPromptInput): AsyncIterable<MicrosoftActivityEvent>;
  cancel(input: CancelMicrosoftTurnInput): Promise<void>;
}

export interface MicrosoftAgentAdapterFactory {
  create(config: BridgeConfig): MicrosoftAgentAdapter;
}

export type MicrosoftActivity = Record<string, unknown> & {
  id?: string;
  type?: string;
  text?: string;
  conversation?: {
    id?: string;
    [key: string]: unknown;
  };
  attachments?: unknown[];
  suggestedActions?: {
    actions?: MicrosoftCardAction[];
    [key: string]: unknown;
  };
  channelData?: Record<string, unknown>;
  entities?: MicrosoftEntity[];
  value?: unknown;
  valueType?: string;
  name?: string;
};

export interface MicrosoftCardAction {
  title?: string;
  text?: string;
  value?: unknown;
  type?: string;
  [key: string]: unknown;
}

export interface MicrosoftEntity {
  type?: string;
  name?: string;
  title?: string;
  url?: string;
  citation?: unknown;
  [key: string]: unknown;
}
