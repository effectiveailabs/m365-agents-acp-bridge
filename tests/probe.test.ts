import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { FakeMicrosoftAgentAdapter } from '../src/microsoft/fakeAdapter.js';
import { runProbe } from '../src/diagnostics/probe.js';
import type { StartedHttpBridgeServer } from '../src/server/httpServer.js';
import { startHttpBridgeServer } from '../src/server/httpServer.js';
import { testConfig } from './helpers.js';

describe('probe diagnostics', () => {
  let started: StartedHttpBridgeServer | undefined;
  let tmpDir: string | undefined;

  afterEach(async () => {
    await started?.close();
    started = undefined;
    if (tmpDir) {
      await rm(tmpDir, { recursive: true, force: true });
      tmpDir = undefined;
    }
  });

  it('exercises initialize, session/new, session/prompt, and streamed updates', async () => {
    const fake = new FakeMicrosoftAgentAdapter({
      promptEvents: [
        {
          activity: {
            id: 'probe-activity-1',
            type: 'message',
            text: 'Probe response',
            conversation: { id: 'probe-conversation-1' },
          },
        },
      ],
    });
    started = await startHttpBridgeServer({
      config: testConfig(),
      microsoft: fake,
    });
    tmpDir = await mkdtemp(join(tmpdir(), 'm365-acp-probe-'));
    const transcriptFile = join(tmpDir, 'transcript.json');

    const result = await runProbe({
      url: `${started.url}/acp`,
      accessToken: 'short-lived-token',
      prompt: 'Say hi',
      timeoutMs: 1000,
      transcriptFile,
    });

    expect(result).toMatchObject({
      ok: true,
      sessionId: 'session-1',
      stopReason: 'end_turn',
      finalText: 'Probe response',
    });
    expect(fake.startSessionCalls[0]?.auth.accessToken).toBe('short-lived-token');
    expect(fake.sendPromptCalls[0]?.prompt).toEqual([{ type: 'text', text: 'Say hi' }]);

    const transcript = await readFile(transcriptFile, 'utf8');
    expect(transcript).toContain('session/prompt');
    expect(transcript).toContain('Probe response');
    expect(transcript).not.toContain('short-lived-token');
  });

  it('fails before network calls when no token is available', async () => {
    const result = await runProbe({
      url: 'http://127.0.0.1:3838/acp',
      timeoutMs: 1,
    });

    expect(result.ok).toBe(false);
    expect(result.error).toContain('No Microsoft access token found');
  });
});
