# Live Microsoft Validation

The bridge should be useful to contributors without Microsoft tenant access. Live validation is a compatibility step after local fake, fixture, protocol-simulator, and package smoke tests pass.

## Local Gates First

Run:

```bash
pnpm check
```

This verifies:

- config validation and secret redaction,
- fake Microsoft adapter lifecycle,
- activity mapping and `_meta.microsoft` preservation,
- best-effort cancellation,
- HTTP/SSE transport behavior,
- Microsoft SDK SSE consumption with stubbed `fetch`,
- package tarball contents, importability, and CLI startup.

## Self-Service Feasibility

Current Microsoft documentation makes fully automated live validation unreliable for generic OSS contributors:

- Microsoft 365 Developer Program E5 sandboxes are available only to qualifying members, such as Visual Studio Professional/Enterprise subscribers or members of qualifying programs. They are development/test tenants and may need recreation every 90 days.
- Copilot Studio trials can create agents and use the test chat panel, but Microsoft says trial licenses cannot publish agents.
- Copilot Studio trial environments expire after 30 days and delete agents/data when the environment expires.
- The Microsoft 365 Agents SDK path requires an existing Copilot Studio agent plus either the SDK connection string/direct-connect URL or expanded metadata.

Relevant Microsoft docs:

- [Microsoft 365 Developer Program](https://developer.microsoft.com/en-us/microsoft-365/dev-program)
- [Microsoft 365 Developer Program FAQ](https://learn.microsoft.com/en-us/office/developer-program/microsoft-365-developer-program-faq)
- [Copilot Studio licensing](https://learn.microsoft.com/en-us/microsoft-copilot-studio/billing-licensing)
- [Get access to Copilot Studio](https://learn.microsoft.com/en-us/microsoft-copilot-studio/requirements-licensing-subscriptions)
- [Work with Power Platform environments](https://learn.microsoft.com/en-us/microsoft-copilot-studio/environments-first-run-experience)

## Agents SDK Prerequisites

For a live Agents SDK invocation test, the validating tenant needs:

- a Copilot Studio agent configured for Microsoft 365 Agents SDK invocation,
- connection string/direct-connect URL, or Environment ID + Tenant ID + Schema name,
- an Entra app registration,
- Power Platform API delegated permission `CopilotStudio.Copilots.Invoke`,
- admin consent where required,
- a short-lived delegated Microsoft access token for the invoking user.

Microsoft references:

- [Integrate with web or native apps using Microsoft 365 Agents SDK](https://learn.microsoft.com/en-us/microsoft-copilot-studio/publication-integrate-web-or-native-app-m365-agents-sdk)
- [Power Platform API authentication](https://learn.microsoft.com/en-us/power-platform/admin/programmability-authentication-v2)
- [Copilot Studio Microsoft authentication setup](https://learn.microsoft.com/en-us/microsoft-copilot-studio/guidance/kit-microsoft-authentication)

## Direct Line Boundary

Microsoft documents Direct Line as the fallback when the Microsoft 365 Agents SDK does not support a scenario, including service principal token scenarios. This bridge's v1 target remains the Agents SDK path with delegated user tokens. Direct Line support can be considered later as a separate adapter mode.

## Credential Gate

Ask for real-world access only after all no-credential tests and self-service options are exhausted.

Minimum live validation inputs:

- tenant ID,
- environment ID,
- schema name or direct-connect URL/connection string as a secret ref,
- app registration client ID,
- short-lived delegated Microsoft access token with `CopilotStudio.Copilots.Invoke`,
- confirmation that the agent is published/configured for Agents SDK invocation.

Do not request or store:

- Microsoft refresh tokens,
- tenant admin passwords,
- plaintext client secrets in chat,
- plaintext connection strings in public issues,
- direct-connect URLs in logs.
