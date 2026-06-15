# Contributing

This project bridges ACP clients to Microsoft Copilot Studio / Microsoft 365 Agents SDK agents.

## Development

```bash
pnpm install
pnpm check
```

`pnpm check` runs linting, formatting checks, TypeScript, unit/integration tests, and a package tarball smoke test.

## Test Philosophy

Most contributions should be testable without a Microsoft tenant:

- use `FakeMicrosoftAgentAdapter` for bridge behavior,
- use Microsoft-like activity fixtures for mapper behavior,
- stub `fetch` for Microsoft SDK streaming behavior,
- keep live tenant assumptions out of unit tests.

Live Microsoft validation is useful, but it must be documented as optional compatibility testing unless the change explicitly affects live Microsoft authentication or invocation.

## Secrets

Do not commit tokens, refresh tokens, client secrets, direct-connect URLs, connection strings, tenant-private URLs, or customer payloads.

Configuration examples must use secret refs:

```json
{
  "secretRef": "env:COPILOT_STUDIO_DIRECT_CONNECT_URL"
}
```

## Scope

Keep the bridge vendor-neutral at runtime. EffectiveAI may consume the package, but EffectiveAI-specific discovery, catalog, refresh-token storage, or customer registration logic belongs outside this repository.
