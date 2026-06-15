# M365 Agents ACP Bridge

Expose Microsoft Copilot Studio / Microsoft 365 Agents SDK agents as Agent Client Protocol-compatible agents.

This package is an ACP server. It accepts ACP requests from any compatible client and invokes Microsoft Copilot Studio underneath through `@microsoft/agents-copilotstudio-client`.

This is not GitHub Copilot. It is for Microsoft enterprise Copilot Studio / M365 Copilot agents.

## Status

Early skeleton. The first implementation focuses on stable configuration, auth contracts, session mapping, activity mapping, and fake-adapter tests. Live Microsoft tenant integration is intentionally deferred until the contracts are stable.

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
m365-agents-acp-bridge serve --config ./m365-agents-acp-bridge.config.json
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
m365-agents-acp-bridge stdio --config ./m365-agents-acp-bridge.config.json
```

HTTP/SSE remains the primary runtime transport for hosted bridge deployments.

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

## Testing

The package is designed to be tested without Microsoft tenant credentials.

Current automated coverage:

- config parsing, auth mode validation, and secret redaction
- Microsoft activity mapping with rich payload preservation under `_meta.microsoft`
- ACP initialize/authenticate/session lifecycle with a fake Microsoft adapter
- best-effort `session/cancel`
- HTTP/SSE JSON-RPC transport smoke test
- stdio ACP transport smoke test through the official ACP TypeScript SDK
- real adapter consumption of Microsoft-like SSE by stubbing `fetch`

Live Microsoft validation is a separate compatibility step. A Copilot Studio trial can create agents and use the test chat panel, but Microsoft documentation says trial licenses cannot publish agents. Agents SDK validation needs an existing Copilot Studio agent, connection string or expanded metadata, an Entra app registration, and delegated `CopilotStudio.Copilots.Invoke` consent.

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
