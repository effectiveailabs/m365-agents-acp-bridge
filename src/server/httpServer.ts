import { createServer, type IncomingMessage, type Server, type ServerResponse } from 'node:http';
import type { AddressInfo } from 'node:net';
import type { AnyMessage } from '@agentclientprotocol/sdk';
import type { BridgeConfig } from '../config/types.js';
import { asBridgeError, BridgeError } from '../errors.js';
import { consoleLogger, type Logger } from '../logging/logger.js';
import type { MicrosoftAgentAdapter } from '../microsoft/types.js';
import { BridgeAgent } from '../acp/bridgeAgent.js';
import type { AcpUpdateSink } from '../acp/updateSink.js';

export interface HttpBridgeServerOptions {
  config: BridgeConfig;
  microsoft: MicrosoftAgentAdapter;
  logger?: Logger;
}

export interface StartedHttpBridgeServer {
  server: Server;
  url: string;
  close(): Promise<void>;
}

interface SseClient {
  id: number;
  response: ServerResponse;
}

export async function startHttpBridgeServer(
  options: HttpBridgeServerOptions,
): Promise<StartedHttpBridgeServer> {
  const logger = options.logger ?? consoleLogger;
  const broker = new SseBroker();
  const agent = new BridgeAgent({
    config: options.config,
    microsoft: options.microsoft,
    updates: broker,
    logger,
  });

  const server = createServer(async (request, response) => {
    try {
      await routeRequest(request, response, options.config, agent, broker);
    } catch (error) {
      const bridgeError = asBridgeError(error, 'MS_INVOKE_FAILED');
      writeJson(response, bridgeError.status, {
        error: {
          code: bridgeError.code,
          message: bridgeError.message,
          data: bridgeError.data,
        },
      });
    }
  });

  await new Promise<void>((resolve) => {
    server.listen(options.config.server.port, options.config.server.host, resolve);
  });

  const address = server.address() as AddressInfo;
  const url = `http://${address.address}:${address.port}`;
  logger.info('started M365 Agents ACP Bridge', {
    endpoint: `${url}/acp`,
    health: `${url}/healthz`,
    mode: options.config.auth.mode,
    agents: options.config.agents.length,
  });

  return {
    server,
    url,
    close() {
      return new Promise((resolve, reject) => {
        server.closeAllConnections();
        server.close((error) => (error ? reject(error) : resolve()));
      });
    },
  };
}

class SseBroker implements AcpUpdateSink {
  private readonly clients = new Map<number, SseClient>();
  private nextId = 1;

  add(response: ServerResponse): void {
    const id = this.nextId++;
    this.clients.set(id, { id, response });
    response.writeHead(200, {
      'Content-Type': 'text/event-stream',
      'Cache-Control': 'no-cache, no-transform',
      Connection: 'keep-alive',
      'X-Accel-Buffering': 'no',
    });
    response.write(': connected\n\n');
    response.on('close', () => {
      this.clients.delete(id);
    });
  }

  async sessionUpdate(update: Parameters<AcpUpdateSink['sessionUpdate']>[0]): Promise<void> {
    this.publish({
      jsonrpc: '2.0',
      method: 'session/update',
      params: update,
    });
  }

  publish(message: AnyMessage): void {
    const data = `event: message\ndata: ${JSON.stringify(message)}\n\n`;
    for (const client of this.clients.values()) {
      client.response.write(data);
    }
  }
}

async function routeRequest(
  request: IncomingMessage,
  response: ServerResponse,
  config: BridgeConfig,
  agent: BridgeAgent,
  broker: SseBroker,
): Promise<void> {
  const url = new URL(request.url ?? '/', `http://${request.headers.host ?? '127.0.0.1'}`);

  if (request.method === 'GET' && url.pathname === '/healthz') {
    writeJson(response, 200, { ok: true });
    return;
  }

  if (request.method === 'GET' && url.pathname === '/readyz') {
    writeJson(response, 200, { ok: true, agents: config.agents.length });
    return;
  }

  if (request.method === 'GET' && isAcpPath(url.pathname)) {
    assertAllowedAcpPath(url.pathname, config);
    broker.add(response);
    return;
  }

  if (request.method === 'DELETE' && isAcpPath(url.pathname)) {
    assertAllowedAcpPath(url.pathname, config);
    response.writeHead(204);
    response.end();
    return;
  }

  if (request.method === 'POST' && isAcpPath(url.pathname)) {
    assertAllowedAcpPath(url.pathname, config);
    const agentId = agentIdFromPath(url.pathname);
    const message = (await readJson(request)) as AnyMessage;
    const auth = bearerToken(request);
    const responseMessage = await dispatchJsonRpc(agent, message, auth, agentId);
    if (responseMessage) {
      writeJson(response, 200, responseMessage);
    } else {
      response.writeHead(202);
      response.end();
    }
    return;
  }

  writeJson(response, 404, { error: { code: 'NOT_FOUND', message: 'Not found' } });
}

async function dispatchJsonRpc(
  agent: BridgeAgent,
  message: AnyMessage,
  bearerAccessToken?: string,
  agentId?: string,
): Promise<AnyMessage | undefined> {
  if (!('method' in message)) {
    throw new BridgeError(
      'MS_INVOKE_FAILED',
      'ACP responses are not accepted by this endpoint',
      400,
    );
  }

  if (!('id' in message)) {
    await dispatchNotification(agent, message.method, message.params);
    return undefined;
  }

  try {
    const result = await dispatchRequest(
      agent,
      message.method,
      message.params,
      bearerAccessToken,
      agentId,
    );
    return {
      jsonrpc: '2.0',
      id: message.id,
      result,
    };
  } catch (error) {
    const bridgeError = asBridgeError(error, 'MS_INVOKE_FAILED');
    return {
      jsonrpc: '2.0',
      id: message.id,
      error: {
        code: bridgeError.status >= 500 ? -32603 : -32602,
        message: bridgeError.message,
        data: {
          code: bridgeError.code,
          ...bridgeError.data,
        },
      },
    };
  }
}

async function dispatchRequest(
  agent: BridgeAgent,
  method: string,
  params: unknown,
  bearerAccessToken?: string,
  agentId?: string,
): Promise<unknown> {
  switch (method) {
    case 'initialize':
      return agent.initialize(params as never);
    case 'authenticate':
      return agent.authenticate(withAuthMeta(params, bearerAccessToken) as never);
    case 'session/new':
      return agent.newSession(withAuthMeta(params, bearerAccessToken, { agentId }) as never);
    case 'session/prompt':
      return agent.prompt(withAuthMeta(params, bearerAccessToken) as never);
    default:
      throw new BridgeError('MS_UNSUPPORTED_ACTIVITY', `Unsupported ACP method: ${method}`, 400);
  }
}

async function dispatchNotification(
  agent: BridgeAgent,
  method: string,
  params: unknown,
): Promise<void> {
  switch (method) {
    case 'session/cancel':
      await agent.cancel(params as never);
      return;
    default:
      return;
  }
}

function isAcpPath(pathname: string): boolean {
  return pathname === '/acp' || /^\/agents\/[^/]+\/acp$/.test(pathname);
}

function assertAllowedAcpPath(pathname: string, config: BridgeConfig): void {
  if (pathname !== '/acp' || config.agents.length === 1) {
    return;
  }

  throw new BridgeError(
    'MS_AGENT_NOT_CONFIGURED',
    'Use /agents/:agentId/acp when multiple agents are configured',
    400,
  );
}

function agentIdFromPath(pathname: string): string | undefined {
  const match = /^\/agents\/([^/]+)\/acp$/.exec(pathname);
  return match ? decodeURIComponent(match[1] ?? '') : undefined;
}

function bearerToken(request: IncomingMessage): string | undefined {
  const header = request.headers.authorization;
  if (!header?.toLowerCase().startsWith('bearer ')) {
    return undefined;
  }
  return header.slice('bearer '.length).trim();
}

function meta(params: unknown): Record<string, unknown> {
  if (!isRecord(params) || !isRecord(params._meta)) {
    return {};
  }
  return params._meta;
}

function withAuthMeta(
  params: unknown,
  bearerAccessToken: string | undefined,
  extraMeta: Record<string, unknown> = {},
): Record<string, unknown> {
  const base = isRecord(params) ? params : {};
  return {
    ...base,
    _meta: {
      ...meta(params),
      ...extraMeta,
      ...(bearerAccessToken ? { accessToken: bearerAccessToken } : {}),
    },
  };
}

async function readJson(request: IncomingMessage): Promise<unknown> {
  const chunks: Buffer[] = [];
  for await (const chunk of request) {
    chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
  }
  return JSON.parse(Buffer.concat(chunks).toString('utf8'));
}

function writeJson(response: ServerResponse, status: number, body: unknown): void {
  response.writeHead(status, {
    'Content-Type': 'application/json',
  });
  response.end(JSON.stringify(body));
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
