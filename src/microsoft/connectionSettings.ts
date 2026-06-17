import { ConnectionSettings } from '@microsoft/agents-copilotstudio-client';
import { resolveSecretRef } from '../config/secrets.js';
import type { CopilotStudioAgentConfig } from '../config/types.js';

export async function connectionSettingsForCopilot(
  copilot: CopilotStudioAgentConfig,
  env: NodeJS.ProcessEnv = process.env,
): Promise<ConnectionSettings> {
  const directConnectUrl = copilot.directConnectUrl
    ? await resolveSecretRef(copilot.directConnectUrl, env)
    : undefined;

  const connectionString = copilot.connectionString
    ? await resolveSecretRef(copilot.connectionString, env)
    : undefined;

  return new ConnectionSettings({
    tenantId: copilot.tenantId,
    appClientId: copilot.clientId,
    environmentId: copilot.environmentId,
    schemaName: copilot.schemaName,
    agentIdentifier: copilot.agentIdentifier,
    directConnectUrl: directConnectUrl ?? connectionString,
    cloud: copilot.cloud as ConstructorParameters<typeof ConnectionSettings>[0]['cloud'],
    customPowerPlatformCloud: copilot.customPowerPlatformCloud,
    useExperimentalEndpoint: copilot.useExperimentalEndpoint,
  });
}
