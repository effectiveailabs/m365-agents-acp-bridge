import type { SessionNotification } from '@agentclientprotocol/sdk';

export interface AcpUpdateSink {
  sessionUpdate(update: SessionNotification): Promise<void>;
}

export class RecordingUpdateSink implements AcpUpdateSink {
  readonly updates: SessionNotification[] = [];

  async sessionUpdate(update: SessionNotification): Promise<void> {
    this.updates.push(update);
  }
}
