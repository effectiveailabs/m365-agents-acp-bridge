import { describe, expect, it } from 'vitest';
import { formatDoctorResult, runDoctor } from '../src/diagnostics/doctor.js';
import { testConfig } from './helpers.js';

const NOW = new Date('2026-01-01T00:00:00.000Z');
const VALID_DIRECT_URL =
  'https://copilotstudio.microsoft.com/environments/default/bots/secret-bot/directline/token';

describe('doctor diagnostics', () => {
  it('accepts delegated Power Platform tokens with CopilotStudio.Copilots.Invoke', async () => {
    const result = await runDoctor({
      config: testConfig(),
      accessToken: jwt({
        aud: 'https://api.powerplatform.com',
        exp: 2_000_000_000,
        scp: 'openid profile CopilotStudio.Copilots.Invoke',
        upn: 'maker@example.test',
      }),
      env: {
        COPILOT_STUDIO_DIRECT_CONNECT_URL: VALID_DIRECT_URL,
      },
      now: NOW,
    });

    expect(result.ok).toBe(true);
    expect(result.checks).toContainEqual({
      severity: 'ok',
      name: 'token scope',
      message: 'CopilotStudio.Copilots.Invoke',
    });

    const formatted = formatDoctorResult(result);
    expect(formatted).toContain('copilotstudio.microsoft.com');
    expect(formatted).not.toContain('secret-bot');
  });

  it('rejects app-only role tokens for the v1 delegated path', async () => {
    const result = await runDoctor({
      config: testConfig(),
      accessToken: jwt({
        aud: 'https://api.powerplatform.com',
        exp: 2_000_000_000,
        roles: ['CopilotStudio.Copilots.Invoke'],
      }),
      env: {
        COPILOT_STUDIO_DIRECT_CONNECT_URL: VALID_DIRECT_URL,
      },
      now: NOW,
    });

    expect(result.ok).toBe(false);
    expect(result.checks).toContainEqual({
      severity: 'fail',
      name: 'token scope',
      message:
        'Found app-only roles:CopilotStudio.Copilots.Invoke; v1 requires delegated scp:CopilotStudio.Copilots.Invoke',
    });
  });

  it('fails when a configured secret ref is unavailable', async () => {
    const result = await runDoctor({
      config: testConfig(),
      accessToken: jwt({
        aud: 'https://api.powerplatform.com',
        exp: 2_000_000_000,
        scp: 'CopilotStudio.Copilots.Invoke',
      }),
      env: {},
      now: NOW,
    });

    expect(result.ok).toBe(false);
    expect(result.checks).toContainEqual({
      severity: 'fail',
      name: 'agent endpoint',
      message: 'Environment secret ref COPILOT_STUDIO_DIRECT_CONNECT_URL is not set',
    });
  });
});

function jwt(payload: Record<string, unknown>): string {
  return ['e30', Buffer.from(JSON.stringify(payload)).toString('base64url'), 'sig'].join('.');
}
