# M365 Agents ACP Bridge

Expose Microsoft Copilot Studio / Microsoft 365 Agents SDK agents as Agent Client Protocol-compatible agents.

This package is an ACP server. It accepts ACP requests from any compatible client and invokes Microsoft Copilot Studio underneath through `@microsoft/agents-copilotstudio-client`.

This is not GitHub Copilot. It is for Microsoft enterprise Copilot Studio / M365 Copilot agents.

## Status

Early implementation. The package has working HTTP/SSE and stdio ACP transports, fake-adapter tests, Microsoft SDK SSE simulator tests, and a published-agent live validation against Copilot Studio / Microsoft 365 Agents SDK.

## Goals

- Run an ACP server for Microsoft Copilot Studio agents.
- Support HTTP/SSE as the primary runtime transport.
- Accept short-lived delegated Microsoft access tokens in `external_token` mode.
- Preserve Microsoft-specific payloads under `_meta.microsoft`.
- Avoid advertising filesystem, terminal, or diff capabilities for Microsoft agents.
- Provide deterministic fakes for tests and local development.

## Non-Goals

- No Microsoft tenant-wide agent discovery.
- No EffectiveAI-specific registration logic.
- No GitHub Copilot support.
- No refresh-token storage in the bridge for `external_token`.
- No filesystem, terminal, diff, or coding-agent permission workflows.

## Install

```bash
pnpm add @effectiveai/m365-agents-acp-bridge
```

## Run

```bash
m365-acp serve --config ./m365-acp.config.json
```

Default endpoint:

```text
http://127.0.0.1:3838/acp
```

The v1 runtime transport is HTTP/SSE:

```text
GET  /healthz
GET  /readyz
POST /acp
GET  /acp
DELETE /acp
POST /agents/:agentId/acp
GET  /agents/:agentId/acp
DELETE /agents/:agentId/acp
```

For local ACP clients that expect stdio NDJSON:

```bash
m365-acp stdio --config ./m365-acp.config.json
```

HTTP/SSE remains the primary runtime transport for hosted bridge deployments.

## Validate Setup

Run `doctor` before starting the bridge. It checks config resolution and the shape of the
delegated Microsoft token without printing access tokens, connection strings, or direct-connect
URLs.

```bash
export MICROSOFT_ACCESS_TOKEN='<delegated-token-with-CopilotStudio.Copilots.Invoke>'
m365-acp doctor --config ./m365-acp.config.json
```

Use `--token-env NAME` if your token lives in a different environment variable.

After the bridge is running, use `probe` for an end-to-end ACP HTTP/SSE smoke test:

```bash
m365-acp probe \
  --url http://127.0.0.1:3838/acp \
  --prompt "Perform live validation." \
  --transcript /tmp/m365-acp-probe.json
```

`probe` opens the SSE stream, sends `initialize`, `session/new`, and `session/prompt`, waits for a
streamed `session/update`, and writes a sanitized transcript when `--transcript` is provided. The
transcript excludes authorization headers and bearer tokens.

## Use From An ACP Client

The bridge exposes one Microsoft Copilot Studio agent as an ACP agent. Your ACP client is
responsible for obtaining a short-lived delegated Microsoft access token and passing it to
the bridge. The bridge does not mint or store Microsoft refresh tokens.

For HTTP/SSE clients:

1. Open `GET /acp` and keep the SSE stream open for `session/update` events.
2. Send JSON-RPC requests to `POST /acp`.
3. Put the delegated Microsoft token on requests that invoke Microsoft:
   `Authorization: Bearer <microsoft-access-token>`.
4. Use `/agents/:agentId/acp` instead of `/acp` when the config contains multiple agents.

Minimal HTTP/SSE flow:

```bash
BASE_URL=http://127.0.0.1:3838
MICROSOFT_ACCESS_TOKEN='<delegated-access-token-with-CopilotStudio.Copilots.Invoke>'

# Terminal 1: receive ACP session/update notifications.
curl -N "$BASE_URL/acp"

# Terminal 2: initialize the ACP agent.
curl -s "$BASE_URL/acp" \
  -H 'content-type: application/json' \
  -d '{
    "jsonrpc": "2.0",
    "id": 1,
    "method": "initialize",
    "params": {
      "protocolVersion": 1,
      "clientCapabilities": {
        "fs": false,
        "terminal": false
      }
    }
  }'

# Create a session. The response contains result.sessionId.
curl -s "$BASE_URL/acp" \
  -H 'content-type: application/json' \
  -H "authorization: Bearer $MICROSOFT_ACCESS_TOKEN" \
  -d '{
    "jsonrpc": "2.0",
    "id": 2,
    "method": "session/new",
    "params": {
      "cwd": "/tmp",
      "mcpServers": []
    }
  }'

# Prompt the Microsoft agent. Read streamed content from Terminal 1.
curl -s "$BASE_URL/acp" \
  -H 'content-type: application/json' \
  -H "authorization: Bearer $MICROSOFT_ACCESS_TOKEN" \
  -d '{
    "jsonrpc": "2.0",
    "id": 3,
    "method": "session/prompt",
    "params": {
      "sessionId": "session-1",
      "prompt": [
        {
          "type": "text",
          "text": "Hello"
        }
      ]
    }
  }'
```

The SSE stream emits ACP notifications like:

```json
{
  "jsonrpc": "2.0",
  "method": "session/update",
  "params": {
    "sessionId": "session-1",
    "update": {
      "sessionUpdate": "agent_message_chunk",
      "content": {
        "type": "text",
        "text": "Hello from Copilot Studio",
        "_meta": {
          "microsoft": {
            "activity": {}
          }
        }
      }
    }
  }
}
```

Cancellation is an ACP notification, so send it without an `id`:

```bash
curl -s "$BASE_URL/acp" \
  -H 'content-type: application/json' \
  -d '{
    "jsonrpc": "2.0",
    "method": "session/cancel",
    "params": {
      "sessionId": "session-1"
    }
  }'
```

Microsoft Copilot Studio agents do not receive filesystem, terminal, or diff tools from this
bridge. ACP clients should advertise those capabilities as unavailable for bridge sessions.

## Config

Secrets are references, not plaintext config values.

```json
{
  "server": {
    "host": "127.0.0.1",
    "port": 3838
  },
  "auth": {
    "mode": "external_token",
    "tokenSource": "authorization_header"
  },
  "agents": [
    {
      "id": "sales-assistant",
      "displayName": "Sales Assistant",
      "copilotStudio": {
        "directConnectUrl": {
          "secretRef": "env:COPILOT_STUDIO_DIRECT_CONNECT_URL"
        }
      }
    }
  ]
}
```

In `external_token` mode the bridge consumes short-lived delegated Microsoft access tokens. The hosting client remains responsible for Microsoft refresh tokens and consent.

Required Microsoft delegated permission:

```text
Power Platform API: CopilotStudio.Copilots.Invoke
```

Optional runtime guard:

```text
M365_ACP_MICROSOFT_STREAM_TIMEOUT_MS=60000
```

This bounds each Microsoft SDK stream read so failed or stalled upstream streams cannot hang ACP requests indefinitely.

## Testing

The package is designed to be tested without Microsoft tenant credentials.

Current automated coverage:

- config parsing, auth mode validation, and secret redaction
- `m365-acp doctor` token/config diagnostics
- `m365-acp probe` HTTP/SSE lifecycle validation
- Microsoft activity mapping with rich payload preservation under `_meta.microsoft`
- ACP initialize/authenticate/session lifecycle with a fake Microsoft adapter
- best-effort `session/cancel`
- HTTP/SSE JSON-RPC transport smoke test
- stdio ACP transport smoke test through the official ACP TypeScript SDK
- real adapter consumption of Microsoft-like SSE by stubbing `fetch`
- Microsoft SDK stalled-stream timeout behavior

Live Microsoft validation is a separate compatibility step. A Copilot Studio trial can create agents and expose a Native app / Agents SDK connection string, but Microsoft documentation says trial licenses cannot publish agents. Pay-as-you-go covers runtime billing, but the maker still needs tenant-level Copilot Studio author access or a qualifying license before they can publish. Agents SDK validation needs an existing published Copilot Studio agent, connection string or expanded metadata, an Entra app registration, delegated `CopilotStudio.Copilots.Invoke` consent, and confirmation that the maker can publish agents.

The current bridge has been validated end to end against a published Copilot Studio agent over ACP HTTP/SSE. The successful token was a delegated Power Platform token with `scp: CopilotStudio.Copilots.Invoke`; app-only `roles: CopilotStudio.Copilots.Invoke` did not produce a usable Agents SDK stream in this test. A delegated token minted through an Entra OBO exchange worked when passed to the bridge as `external_token`.

See [docs/live-validation.md](docs/live-validation.md) for the live Microsoft validation ladder and credential gate.

## Development

```bash
pnpm install
pnpm check
pnpm test
pnpm build
```

`pnpm check` also runs a package tarball smoke test to confirm built output imports cleanly, the CLI starts, and test secret fixtures are not published.

## License

Apache-2.0
