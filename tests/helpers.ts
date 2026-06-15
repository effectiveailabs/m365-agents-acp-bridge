import type { BridgeConfig } from '../src/config/types.js';

export function testConfig(overrides: Partial<BridgeConfig> = {}): BridgeConfig {
  return {
    server: {
      host: '127.0.0.1',
      port: 0,
      ...overrides.server,
    },
    auth: {
      mode: 'external_token',
      tokenSource: 'authorization_header',
      ...overrides.auth,
    },
    agents: overrides.agents ?? [
      {
        id: 'default',
        displayName: 'Default Copilot Studio Agent',
        copilotStudio: {
          directConnectUrl: {
            secretRef: 'env:COPILOT_STUDIO_DIRECT_CONNECT_URL',
          },
        },
      },
    ],
    logLevel: overrides.logLevel,
  };
}

export function wait(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
