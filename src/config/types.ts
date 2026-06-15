export type AuthMode = 'external_token' | 'interactive' | 'obo';

export type TokenSource = 'authorization_header' | 'authenticate_meta' | 'env';

export interface SecretRef {
  secretRef: string;
}

export interface ServerConfig {
  host: string;
  port: number;
}

export interface AuthConfig {
  mode: AuthMode;
  tokenSource?: TokenSource;
}

export interface CopilotStudioAgentConfig {
  tenantId?: string;
  environmentId?: string;
  schemaName?: string;
  agentIdentifier?: string;
  clientId?: string;
  clientSecret?: SecretRef;
  connectionString?: SecretRef;
  directConnectUrl?: SecretRef;
  cloud?: string;
  customPowerPlatformCloud?: string;
  useExperimentalEndpoint?: boolean;
}

export interface AgentConfig {
  id: string;
  displayName?: string;
  description?: string;
  copilotStudio: CopilotStudioAgentConfig;
}

export interface BridgeConfig {
  server: ServerConfig;
  auth: AuthConfig;
  agents: AgentConfig[];
  logLevel?: 'debug' | 'info' | 'warn' | 'error';
}

export interface LoadConfigOptions {
  configFile?: string;
  env?: NodeJS.ProcessEnv;
}
