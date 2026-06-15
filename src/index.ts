export { BridgeAgent } from './acp/bridgeAgent.js';
export type { BridgeAgentOptions } from './acp/bridgeAgent.js';
export { RecordingUpdateSink } from './acp/updateSink.js';
export type { AcpUpdateSink } from './acp/updateSink.js';
export {
  loadBridgeConfig,
  parseBridgeConfig,
  safeConfigForLog,
  validateBridgeConfig,
} from './config/load.js';
export { redactSensitive, redactString } from './config/redact.js';
export { resolveSecretRef, SecretRefError } from './config/secrets.js';
export type {
  AgentConfig,
  AuthConfig,
  AuthMode,
  BridgeConfig,
  CopilotStudioAgentConfig,
  LoadConfigOptions,
  SecretRef,
  ServerConfig,
  TokenSource,
} from './config/types.js';
export { BridgeError } from './errors.js';
export type { BridgeErrorCode } from './errors.js';
export { FakeMicrosoftAgentAdapter } from './microsoft/fakeAdapter.js';
export type { FakeAdapterScript, FakePromptEvent } from './microsoft/fakeAdapter.js';
export { mapMicrosoftActivityToAcpUpdate } from './microsoft/activityMapper.js';
export { CopilotStudioMicrosoftAdapter } from './microsoft/realAdapter.js';
export type {
  CancelMicrosoftTurnInput,
  MicrosoftActivity,
  MicrosoftActivityEvent,
  MicrosoftAgentAdapter,
  MicrosoftAgentAdapterFactory,
  MicrosoftSession,
  RequestAuthContext,
  SendMicrosoftPromptInput,
  StartMicrosoftSessionInput,
} from './microsoft/types.js';
export { MemorySessionStore } from './session/memorySessionStore.js';
export type { BridgeSessionRecord, SessionStore } from './session/memorySessionStore.js';
export { startHttpBridgeServer } from './server/httpServer.js';
export type { HttpBridgeServerOptions, StartedHttpBridgeServer } from './server/httpServer.js';
