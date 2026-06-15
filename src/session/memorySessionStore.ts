import type { MicrosoftSession } from '../microsoft/types.js';

export interface BridgeSessionRecord {
  sessionId: string;
  agentId: string;
  microsoft: MicrosoftSession;
  createdAt: Date;
  updatedAt: Date;
}

export interface SessionStore {
  create(input: { agentId: string; microsoft: MicrosoftSession }): Promise<BridgeSessionRecord>;
  get(sessionId: string): Promise<BridgeSessionRecord | undefined>;
  updateMicrosoft(sessionId: string, microsoft: MicrosoftSession): Promise<BridgeSessionRecord>;
  delete(sessionId: string): Promise<void>;
}

export class MemorySessionStore implements SessionStore {
  private readonly sessions = new Map<string, BridgeSessionRecord>();
  private nextSessionNumber = 1;

  async create(input: {
    agentId: string;
    microsoft: MicrosoftSession;
  }): Promise<BridgeSessionRecord> {
    const now = new Date();
    const record: BridgeSessionRecord = {
      sessionId: `session-${this.nextSessionNumber++}`,
      agentId: input.agentId,
      microsoft: input.microsoft,
      createdAt: now,
      updatedAt: now,
    };
    this.sessions.set(record.sessionId, record);
    return record;
  }

  async get(sessionId: string): Promise<BridgeSessionRecord | undefined> {
    return this.sessions.get(sessionId);
  }

  async updateMicrosoft(
    sessionId: string,
    microsoft: MicrosoftSession,
  ): Promise<BridgeSessionRecord> {
    const existing = this.sessions.get(sessionId);
    if (!existing) {
      throw new Error(`Session not found: ${sessionId}`);
    }

    const updated: BridgeSessionRecord = {
      ...existing,
      microsoft: {
        ...existing.microsoft,
        ...microsoft,
      },
      updatedAt: new Date(),
    };
    this.sessions.set(sessionId, updated);
    return updated;
  }

  async delete(sessionId: string): Promise<void> {
    this.sessions.delete(sessionId);
  }
}
