import { writeFile } from 'node:fs/promises';

export interface ProbeOptions {
  url: string;
  accessToken?: string;
  prompt?: string;
  timeoutMs?: number;
  transcriptFile?: string;
  fetch?: typeof fetch;
}

export interface ProbeResult {
  ok: boolean;
  sessionId?: string;
  stopReason?: string;
  finalText?: string;
  updates: unknown[];
  error?: string;
  transcriptFile?: string;
}

interface JsonRpcResponse {
  jsonrpc?: string;
  id?: string | number | null;
  result?: unknown;
  error?: {
    code?: unknown;
    message?: unknown;
    data?: unknown;
  };
}

interface TranscriptEntry {
  direction: 'request' | 'response' | 'sse';
  method?: string;
  id?: number;
  status?: number;
  body?: unknown;
}

const DEFAULT_PROBE_PROMPT = 'Perform live validation.';
const DEFAULT_PROBE_TIMEOUT_MS = 30_000;

export async function runProbe(options: ProbeOptions): Promise<ProbeResult> {
  const fetchImpl = options.fetch ?? fetch;
  const timeoutMs = options.timeoutMs ?? DEFAULT_PROBE_TIMEOUT_MS;
  const prompt = options.prompt ?? DEFAULT_PROBE_PROMPT;
  const endpoint = acpEndpoint(options.url);
  const transcript: TranscriptEntry[] = [];
  const updates: unknown[] = [];

  if (!options.accessToken) {
    return {
      ok: false,
      updates,
      error:
        'No Microsoft access token found. Set MICROSOFT_ACCESS_TOKEN or M365_ACP_MICROSOFT_ACCESS_TOKEN.',
    };
  }

  let reader: ReadableStreamDefaultReader<Uint8Array> | undefined;
  const streamController = new AbortController();

  try {
    const stream = await fetchImpl(endpoint, {
      method: 'GET',
      headers: { accept: 'text/event-stream' },
      signal: streamController.signal,
    });
    if (!stream.ok || !stream.body) {
      throw new Error(`SSE stream failed: HTTP ${stream.status}`);
    }

    reader = stream.body.getReader();
    const sse = new SseJsonReader(reader);

    await postRpc(fetchImpl, endpoint, options.accessToken, transcript, {
      jsonrpc: '2.0',
      id: 1,
      method: 'initialize',
      params: {
        protocolVersion: 1,
        clientCapabilities: {
          fs: false,
          terminal: false,
        },
      },
    });

    const sessionResponse = await postRpc(fetchImpl, endpoint, options.accessToken, transcript, {
      jsonrpc: '2.0',
      id: 2,
      method: 'session/new',
      params: {
        cwd: '/tmp',
        mcpServers: [],
      },
    });
    const sessionId = sessionIdFromResult(sessionResponse.result);

    const promptResponsePromise = postRpc(fetchImpl, endpoint, options.accessToken, transcript, {
      jsonrpc: '2.0',
      id: 3,
      method: 'session/prompt',
      params: {
        sessionId,
        prompt: [{ type: 'text', text: prompt }],
      },
    });

    const update = await readUntilTextUpdate(sse, sessionId, updates, transcript, timeoutMs);
    const promptResponse = await promptResponsePromise;
    const stopReason = stopReasonFromResult(promptResponse.result);
    const finalText = textFromUpdate(update);

    const result: ProbeResult = {
      ok: true,
      sessionId,
      stopReason,
      finalText,
      updates,
    };
    await writeTranscript(options.transcriptFile, endpoint, transcript, result);
    return { ...result, transcriptFile: options.transcriptFile };
  } catch (error) {
    const result: ProbeResult = {
      ok: false,
      updates,
      error: errorMessage(error),
    };
    await writeTranscript(options.transcriptFile, endpoint, transcript, result);
    return { ...result, transcriptFile: options.transcriptFile };
  } finally {
    streamController.abort();
    void reader?.cancel().catch(() => undefined);
  }
}

export function formatProbeResult(result: ProbeResult): string {
  const lines = ['M365 ACP probe'];

  if (result.sessionId) {
    lines.push(`Session: ${result.sessionId}`);
  }
  if (result.stopReason) {
    lines.push(`Stop reason: ${result.stopReason}`);
  }
  if (result.finalText) {
    lines.push(`Final text: ${result.finalText}`);
  }
  if (result.transcriptFile) {
    lines.push(`Transcript: ${result.transcriptFile}`);
  }
  if (result.error) {
    lines.push(`Error: ${result.error}`);
  }

  lines.push(result.ok ? 'Result: ok' : 'Result: failed');
  return `${lines.join('\n')}\n`;
}

function acpEndpoint(rawUrl: string): string {
  const url = new URL(rawUrl);
  if (url.pathname === '/' || url.pathname === '') {
    url.pathname = '/acp';
  }
  return url.toString();
}

async function postRpc(
  fetchImpl: typeof fetch,
  endpoint: string,
  accessToken: string,
  transcript: TranscriptEntry[],
  body: {
    jsonrpc: '2.0';
    id: number;
    method: string;
    params: Record<string, unknown>;
  },
): Promise<JsonRpcResponse> {
  transcript.push({
    direction: 'request',
    id: body.id,
    method: body.method,
    body,
  });

  const response = await fetchImpl(endpoint, {
    method: 'POST',
    headers: {
      authorization: `Bearer ${accessToken}`,
      'content-type': 'application/json',
    },
    body: JSON.stringify(body),
  });
  const text = await response.text();
  const parsed = parseJsonResponse(text);

  transcript.push({
    direction: 'response',
    id: body.id,
    method: body.method,
    status: response.status,
    body: parsed,
  });

  if (!response.ok) {
    throw new Error(`HTTP ${response.status}: ${text}`);
  }

  if (parsed.error) {
    throw new Error(String(parsed.error.message ?? 'JSON-RPC error'));
  }

  return parsed;
}

async function readUntilTextUpdate(
  sse: SseJsonReader,
  sessionId: string,
  updates: unknown[],
  transcript: TranscriptEntry[],
  timeoutMs: number,
): Promise<unknown> {
  const deadline = Date.now() + timeoutMs;

  for (;;) {
    const remainingMs = deadline - Date.now();
    if (remainingMs <= 0) {
      throw new Error(`Timed out waiting for ACP session/update after ${timeoutMs}ms`);
    }

    const message = await sse.nextJson(remainingMs);
    transcript.push({ direction: 'sse', body: message });

    if (!isRecord(message) || message.method !== 'session/update') {
      continue;
    }

    const params = isRecord(message.params) ? message.params : {};
    const update = params.update;
    updates.push(update);

    if (params.sessionId !== sessionId) {
      continue;
    }

    if (textFromUpdate(update)) {
      return update;
    }
  }
}

class SseJsonReader {
  private readonly decoder = new TextDecoder();
  private buffer = '';

  constructor(private readonly reader: ReadableStreamDefaultReader<Uint8Array>) {}

  async nextJson(timeoutMs: number): Promise<unknown> {
    for (;;) {
      const parsed = this.readBufferedEvent();
      if (parsed !== undefined) {
        return parsed;
      }

      const read = await readWithTimeout(this.reader, timeoutMs);
      if (read.done) {
        throw new Error('SSE stream closed before session/update');
      }

      this.buffer += this.decoder.decode(read.value, { stream: true });
    }
  }

  private readBufferedEvent(): unknown | undefined {
    for (;;) {
      const separator = this.buffer.indexOf('\n\n');
      if (separator < 0) {
        return undefined;
      }

      const rawEvent = this.buffer.slice(0, separator);
      this.buffer = this.buffer.slice(separator + 2);
      const data = rawEvent
        .split('\n')
        .filter((line) => line.startsWith('data: '))
        .map((line) => line.slice('data: '.length))
        .join('\n');

      if (!data) {
        continue;
      }

      return JSON.parse(data) as unknown;
    }
  }
}

async function readWithTimeout(
  reader: ReadableStreamDefaultReader<Uint8Array>,
  timeoutMs: number,
): Promise<ReadableStreamReadResult<Uint8Array>> {
  let timeout: ReturnType<typeof setTimeout> | undefined;

  try {
    return await Promise.race([
      reader.read(),
      new Promise<never>((_, reject) => {
        timeout = setTimeout(() => {
          reject(new Error(`Timed out waiting for SSE data after ${timeoutMs}ms`));
        }, timeoutMs);
      }),
    ]);
  } finally {
    if (timeout) {
      clearTimeout(timeout);
    }
  }
}

function parseJsonResponse(text: string): JsonRpcResponse {
  try {
    return JSON.parse(text) as JsonRpcResponse;
  } catch {
    throw new Error(`Invalid JSON response: ${text}`);
  }
}

function sessionIdFromResult(result: unknown): string {
  if (isRecord(result) && typeof result.sessionId === 'string') {
    return result.sessionId;
  }
  throw new Error('session/new response did not include result.sessionId');
}

function stopReasonFromResult(result: unknown): string | undefined {
  return isRecord(result) && typeof result.stopReason === 'string' ? result.stopReason : undefined;
}

function textFromUpdate(update: unknown): string | undefined {
  if (!isRecord(update) || !isRecord(update.content)) {
    return undefined;
  }

  const text = update.content.text;
  return typeof text === 'string' && text.length > 0 ? text : undefined;
}

async function writeTranscript(
  transcriptFile: string | undefined,
  endpoint: string,
  entries: TranscriptEntry[],
  result: ProbeResult,
): Promise<void> {
  if (!transcriptFile) {
    return;
  }

  await writeFile(
    transcriptFile,
    `${JSON.stringify(
      {
        endpoint,
        entries,
        result,
      },
      null,
      2,
    )}\n`,
    'utf8',
  );
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
