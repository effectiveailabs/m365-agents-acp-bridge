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
export { accessTokenFromEnv, decodeAccessToken, hasScope } from './auth/token.js';
export type { DecodedAccessToken } from './auth/token.js';
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
export { formatDoctorResult, runDoctor } from './diagnostics/doctor.js';
export type {
  DoctorCheck,
  DoctorOptions,
  DoctorResult,
  DoctorSeverity,
} from './diagnostics/doctor.js';
export { formatProbeResult, runProbe } from './diagnostics/probe.js';
export type { ProbeOptions, ProbeResult } from './diagnostics/probe.js';
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
export { createAcpAgentConnection, runStdioBridgeServer } from './server/stdioServer.js';
export type { StdioBridgeServerOptions } from './server/stdioServer.js';
