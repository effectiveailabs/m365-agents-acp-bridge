import { describe, expect, it } from 'vitest';
import { loadBridgeConfig, parseBridgeConfig, safeConfigForLog } from '../src/config/load.js';
import { redactSensitive } from '../src/config/redact.js';

describe('config', () => {
  it('accepts JSON config with secret refs', () => {
    const config = parseBridgeConfig({
      server: { host: '127.0.0.1', port: 3838 },
      auth: { mode: 'external_token', tokenSource: 'authorization_header' },
      agents: [
        {
          id: 'sales',
          copilotStudio: {
            directConnectUrl: { secretRef: 'env:COPILOT_STUDIO_DIRECT_CONNECT_URL' },
          },
        },
      ],
    });

    expect(config.agents[0]?.id).toBe('sales');
    expect(config.auth.mode).toBe('external_token');
  });

  it('maps env secrets to secret refs without storing plaintext values', async () => {
    const config = await loadBridgeConfig({
      env: {
        M365_ACP_DIRECT_CONNECT_URL: 'https://secret.example/direct',
        M365_ACP_AGENT_ID: 'default',
      },
    });

    expect(config.agents[0]?.copilotStudio.directConnectUrl).toEqual({
      secretRef: 'env:M365_ACP_DIRECT_CONNECT_URL',
    });
    expect(JSON.stringify(config)).not.toContain('secret.example');
  });

  it('rejects plaintext secrets where secret refs are required', () => {
    expect(() =>
      parseBridgeConfig({
        agents: [
          {
            id: 'bad',
            copilotStudio: {
              directConnectUrl: 'https://plaintext.example/direct',
            },
          },
        ],
      }),
    ).toThrow();
  });

  it('rejects obo mode in v1', () => {
    expect(() =>
      parseBridgeConfig({
        auth: { mode: 'obo' },
        agents: [
          {
            id: 'default',
            copilotStudio: {
              directConnectUrl: { secretRef: 'env:COPILOT_STUDIO_DIRECT_CONNECT_URL' },
            },
          },
        ],
      }),
    ).toThrow(/obo/);
  });

  it('redacts sensitive fields from safe log payloads', () => {
    const redacted = redactSensitive({
      authorization: 'Bearer eyJsecret',
      accessToken: 'abcdef123456789',
      clientSecret: 'super-secret',
      directConnectUrl: 'https://secret.example/direct',
      keep: 'visible',
      secretRef: 'env:VISIBLE_REF',
    });

    const serialized = JSON.stringify(redacted);
    expect(serialized).not.toContain('eyJsecret');
    expect(serialized).not.toContain('abcdef123456789');
    expect(serialized).not.toContain('super-secret');
    expect(serialized).not.toContain('secret.example');
    expect(serialized).toContain('visible');
    expect(serialized).toContain('env:VISIBLE_REF');
  });

  it('safeConfigForLog keeps refs but not resolved secret values', () => {
    const config = parseBridgeConfig({
      agents: [
        {
          id: 'default',
          copilotStudio: {
            directConnectUrl: { secretRef: 'env:COPILOT_STUDIO_DIRECT_CONNECT_URL' },
          },
        },
      ],
    });

    expect(JSON.stringify(safeConfigForLog(config))).toContain(
      'env:COPILOT_STUDIO_DIRECT_CONNECT_URL',
    );
  });
});
