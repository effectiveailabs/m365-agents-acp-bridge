import type {
  CancelMicrosoftTurnInput,
  MicrosoftActivity,
  MicrosoftActivityEvent,
  MicrosoftAgentAdapter,
  SendMicrosoftPromptInput,
  StartMicrosoftSessionInput,
} from './types.js';

export interface FakeAdapterScript {
  startSession?: {
    conversationId?: string;
    activityId?: string;
  };
  promptEvents?: FakePromptEvent[];
}

export interface FakePromptEvent {
  activity: MicrosoftActivity;
  delayMs?: number;
}

export class FakeMicrosoftAgentAdapter implements MicrosoftAgentAdapter {
  readonly startSessionCalls: StartMicrosoftSessionInput[] = [];
  readonly sendPromptCalls: SendMicrosoftPromptInput[] = [];
  readonly cancelCalls: CancelMicrosoftTurnInput[] = [];
  private readonly script: FakeAdapterScript;
  private cancelled = false;

  constructor(script: FakeAdapterScript = {}) {
    this.script = script;
  }

  async startSession(input: StartMicrosoftSessionInput) {
    this.startSessionCalls.push(input);
    return {
      conversationId: this.script.startSession?.conversationId ?? 'fake-conversation-1',
      activityId: this.script.startSession?.activityId,
    };
  }

  async *sendPrompt(input: SendMicrosoftPromptInput): AsyncIterable<MicrosoftActivityEvent> {
    this.sendPromptCalls.push(input);
    this.cancelled = false;

    const events = this.script.promptEvents ?? [
      {
        activity: {
          id: 'fake-activity-1',
          type: 'message',
          text: 'Fake Microsoft response',
          conversation: { id: input.session.conversationId ?? 'fake-conversation-1' },
        },
      },
    ];

    for (const event of events) {
      if (this.cancelled || input.signal?.aborted) {
        return;
      }
      if (event.delayMs) {
        await delay(event.delayMs);
      }
      if (this.cancelled || input.signal?.aborted) {
        return;
      }
      yield { activity: event.activity };
    }
  }

  async cancel(input: CancelMicrosoftTurnInput): Promise<void> {
    this.cancelCalls.push(input);
    this.cancelled = true;
  }
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
