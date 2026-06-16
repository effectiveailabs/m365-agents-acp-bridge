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

## CLI Probes Before Requesting Credentials

Maintainers should exhaust read-only tenant probes before asking for live agent access. These commands do not print access tokens or secrets.

Authenticate with tenant-level access when no Azure subscription is available:

```bash
az login --allow-no-subscriptions --use-device-code
az account show --query '{tenantId:tenantId,user:user.name,name:name}' --output json
```

Check that Azure CLI can mint a Power Platform API token without printing it:

```bash
az account get-access-token \
  --resource https://api.powerplatform.com \
  --query '{tenant:tenant,expiresOn:expiresOn,tokenType:tokenType}' \
  --output json
```

List Power Platform environments visible to the signed-in user:

```bash
az rest \
  --method get \
  --resource https://api.powerplatform.com \
  --url 'https://api.powerplatform.com/environmentmanagement/environments?api-version=2024-10-01' \
  --output json
```

An empty `value: []` response means the login works, but the user currently has no visible Power Platform environment through this API. It is not enough for Agents SDK live validation.

Check whether tenant policy blocks non-admin environment creation:

```bash
az rest \
  --method post \
  --resource https://api.bap.microsoft.com \
  --url 'https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/listtenantsettings?api-version=2020-10-01' \
  --query '{disableEnvironmentCreationByNonAdminUsers:disableEnvironmentCreationByNonAdminUsers,disableTrialEnvironmentCreationByNonAdminUsers:disableTrialEnvironmentCreationByNonAdminUsers,disableDeveloperEnvironmentCreationByNonAdminUsers:powerPlatform.governance.disableDeveloperEnvironmentCreationByNonAdminUsers}' \
  --output json
```

Verify that the tenant exposes the Power Platform API service principal and the Copilot Studio invoke delegated scope:

```bash
az ad sp list \
  --display-name "Power Platform API" \
  --query '[].{appId:appId,oauth2PermissionScopes:oauth2PermissionScopes[?value==`CopilotStudio.Copilots.Invoke`].value,appRoles:appRoles[?value==`CopilotStudio.Copilots.Invoke`].value}' \
  --output json
```

The expected API app ID is `8578e004-a5c6-46e7-913e-12f58912df43`. The delegated scope value needed by the bridge is `CopilotStudio.Copilots.Invoke`.

The Power Platform CLI (`pac`) is useful when it authenticates cleanly, but live validation should not depend on it. If `pac auth create --deviceCode` fails locally, prefer the Azure CLI probes above and continue with explicit live agent inputs.

If environment creation is not blocked but the environment list is empty, the remaining setup is a Microsoft product/licensing step: sign up for Power Apps Developer Plan, Power Apps trial, Copilot Studio trial, or use an existing paid tenant/environment. Microsoft documents that Microsoft 365 licenses alone do not allow users to manage environments.

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

## Observed Compatibility Run

On 2026-06-16, a self-service test tenant validated the bridge against the real Copilot Studio / Microsoft 365 Agents SDK endpoint without storing secrets in the repo:

- A tenant-level Azure CLI login worked without an Azure subscription.
- A Copilot Studio trial allowed creating a Developer environment and a test agent.
- Sandbox environment creation failed because the tenant had no available Dataverse database capacity.
- Trial licensing blocked publishing the agent.
- The Native app channel still exposed an Agents SDK direct-connect URL before publish.
- Azure CLI's built-in client could mint a Power Platform token, but that token carried `CopilotStudio.Copilots.Test`, not `CopilotStudio.Copilots.Invoke`; the Agents SDK endpoint returned `403 InsufficientDelegatedPermissions`.
- A tenant-local public-client app registration with delegated Power Platform API `CopilotStudio.Copilots.Invoke` consent produced a short-lived token with `scp: CopilotStudio.Copilots.Invoke`.
- A direct POST to the Agents SDK SSE endpoint returned `200 OK`, `text/event-stream`, and a Microsoft conversation header.
- Because the trial agent could not be published, Microsoft returned an activity text containing `LatestPublishedVersionNotFound`.
- The packaged bridge completed real HTTP/SSE ACP `initialize`, `session/new`, and `session/prompt` calls using the invoke-scoped delegated token, including preserving Microsoft activity data under `_meta.microsoft`.

This proves the ACP transport, delegated-token handoff, Microsoft SDK connection string path, and activity streaming path against the real Microsoft endpoint. It does not prove successful business-agent response generation because the test tenant could not publish the agent.

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
