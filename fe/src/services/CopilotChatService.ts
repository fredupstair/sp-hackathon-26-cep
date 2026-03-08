import { MSGraphClientV3 } from '@microsoft/sp-http';
import { COPILOT_SYSTEM_PROMPT, PERSONALIZED_WELCOME_PROMPT } from './CopilotSystemPrompts';

/**
 * Copilot Conversations API via Microsoft Graph (beta)
 * POST https://graph.microsoft.com/beta/copilot/conversations
 * POST https://graph.microsoft.com/beta/copilot/conversations/{id}/chat
 *
 * Uses MSGraphClientV3 — no extra webApiPermissionRequests needed.
 */
export class CopilotChatService {
  private readonly _graphClient: MSGraphClientV3;

  constructor(graphClient: MSGraphClientV3) {
    this._graphClient = graphClient;
  }

  /** POST /beta/copilot/conversations → returns conversation id */
  async createConversation(): Promise<string> {
    const response = await this._graphClient
      .api('/copilot/conversations')
      .version('beta')
      .post({}) as { id?: string };

    if (!response?.id) throw new Error('Copilot API returned no conversation ID');
    return response.id;
  }

  /** POST /beta/copilot/conversations/{id}/chat → returns assistant text */
  async sendMessage(conversationId: string, text: string): Promise<string> {
    const body = {
      message: { text },
      locationHint: { timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone },
      contextualResources: {
        files: [],
        webContext: { isWebEnabled: false },
      },
    };

    const response = await this._graphClient
      .api(`/copilot/conversations/${conversationId}/chat`)
      .version('beta')
      .post(body) as { messages?: Array<{ text?: string }> };

    // First message is the user prompt, second is the assistant reply
    const assistantText = response?.messages?.[1]?.text ?? '';
    if (!assistantText) throw new Error('No response message received from Copilot');
    return assistantText;
  }

  /**
   * Creates a conversation, sends a combined system + user prompt,
   * and returns the generated welcome text.
   *
   * Falls back to a built-in template when the Copilot API is unavailable
   * (no M365 Copilot licence, tenant not configured, etc.).
   */
  async generateWelcomeText(
    organizationName: string,
    customInstruction?: string
  ): Promise<{ text: string; fromFallback: boolean }> {
    try {
      const conversationId = await this.createConversation();

      const userContext =
        `Organisation: "${organizationName}".` +
        (customInstruction ? ` ${customInstruction}` : ' Write the welcome text now.');

      // Combine system prompt + user context in the single message
      const fullPrompt = `${COPILOT_SYSTEM_PROMPT}\n\n${userContext}`;
      const text = await this.sendMessage(conversationId, fullPrompt);
      return { text, fromFallback: false };
    } catch (err) {
      if (this._isApiUnavailableError(err)) {
        return {
          text: CopilotChatService.generateFallbackText(organizationName),
          fromFallback: true,
        };
      }
      throw err;
    }
  }

  private _isApiUnavailableError(err: unknown): boolean {
    const msg = (err instanceof Error ? err.message : String(err)).toLowerCase();
    return (
      msg.includes('500011') ||
      msg.includes('invalid_resource') ||
      msg.includes('resource principal') ||
      msg.includes('was not found in the tenant') ||
      msg.includes('403') ||
      msg.includes('no response message')
    );
  }

  static generateFallbackText(organizationName: string): string {
    return (
      `Join ${organizationName}'s Copilot Engagement Program and transform how you work every day. ` +
      `By tracking your Microsoft 365 Copilot usage you'll earn points, climb our leaderboard, ` +
      `and gain recognition for your AI-powered productivity. Whether you're drafting emails, ` +
      `analysing data, or summarising meetings — every interaction counts. ` +
      `Join a growing community of colleagues already getting more done with Copilot.`
    );
  }

  /**
   * Generates a short (2–3 sentence) personalised welcome for a specific end user.
   * Uses the user's first name, job title, and department to tailor the message.
   * Falls back to a template when the Copilot API is unavailable.
   */
  async generatePersonalizedText(
    displayName: string,
    jobTitle: string,
    department: string,
    organizationName: string
  ): Promise<{ text: string; fromFallback: boolean }> {
    try {
      const conversationId = await this.createConversation();
      const firstName = displayName.split(' ')[0];
      const roleContext = [jobTitle, department].filter(Boolean).join(', ');
      const userContext =
        `User: ${firstName}${roleContext ? ` (${roleContext})` : ''}. ` +
        `Organisation: "${organizationName || 'our company'}". Write the short welcome message now.`;
      const fullPrompt = `${PERSONALIZED_WELCOME_PROMPT}\n\n${userContext}`;
      const text = await this.sendMessage(conversationId, fullPrompt);
      return { text, fromFallback: false };
    } catch (err) {
      if (this._isApiUnavailableError(err)) {
        return {
          text: CopilotChatService.generatePersonalizedFallbackText(displayName, jobTitle, organizationName),
          fromFallback: true,
        };
      }
      throw err;
    }
  }

  /**
   * Streams a personalised welcome using chatOverStream (SSE).
   * Calls onChunk with the accumulated assistant text as it arrives.
   * Falls back to the non-streaming method if streaming is unavailable.
   */
  async streamPersonalizedText(
    displayName: string,
    jobTitle: string,
    department: string,
    organizationName: string,
    onChunk: (accumulatedText: string) => void,
  ): Promise<{ text: string; fromFallback: boolean }> {
    try {
      const conversationId = await this.createConversation();
      const firstName = displayName.split(' ')[0];
      const roleContext = [jobTitle, department].filter(Boolean).join(', ');
      const userContext =
        `User: ${firstName}${roleContext ? ` (${roleContext})` : ''}. ` +
        `Organisation: "${organizationName || 'our company'}". Write the short welcome message now.`;
      const fullPrompt = `${PERSONALIZED_WELCOME_PROMPT}\n\n${userContext}`;

      const text = await this._sendStreamMessage(conversationId, fullPrompt, onChunk);
      return { text, fromFallback: false };
    } catch {
      // If streaming fails, try non-streaming as fallback
      try {
        const result = await this.generatePersonalizedText(displayName, jobTitle, department, organizationName);
        onChunk(result.text);
        return result;
      } catch {
        const fallbackText = CopilotChatService.generatePersonalizedFallbackText(displayName, jobTitle, organizationName);
        onChunk(fallbackText);
        return { text: fallbackText, fromFallback: true };
      }
    }
  }

  /**
   * POST /beta/copilot/conversations/{id}/chatOverStream
   * Returns a text/event-stream (SSE). Each event is a copilotConversation
   * whose messages array contains the accumulated assistant text.
   */
  private async _sendStreamMessage(
    conversationId: string,
    text: string,
    onChunk: (accumulated: string) => void,
  ): Promise<string> {
    const body = {
      message: { text },
      locationHint: { timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone },
      contextualResources: {
        files: [],
        webContext: { isWebEnabled: false },
      },
    };

    // Use responseType 'raw' to get the native Response for SSE parsing
    const rawResponse: Response = await (this._graphClient
      .api(`/copilot/conversations/${conversationId}/chatOverStream`)
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      .version('beta') as any)
      .responseType('raw')
      .post(body);

    if (!rawResponse.body) {
      throw new Error('No response body from chatOverStream');
    }

    const reader = rawResponse.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';

    // ── Smooth typewriter reveal ──────────────────────────────────────────────
    // SSE events fill targetText; a parallel timer reveals characters at an
    // adaptive pace: slow when the unrevealed buffer is thin (so it never
    // fully catches up during streaming), faster when there is plenty of
    // runway or the stream has finished.
    let targetText = '';        // Full accumulated text from SSE so far
    let revealedLen = 0;        // How many characters have been shown via onChunk
    let streamDone = false;     // True once the SSE read loop finishes
    const BUFFER_CHARS = 120;   // Accumulate this many chars before starting reveal
    const TICK_MS = 25;         // Base interval between reveals

    const revealDone = new Promise<void>((resolve) => {
      const tick = setInterval(() => {
        // Don't start until we have enough runway (or the stream already ended)
        if (targetText.length < BUFFER_CHARS && !streamDone) return;

        const slack = targetText.length - revealedLen;
        let chars: number;

        if (streamDone) {
          // Stream finished — flush remaining text quickly but still smoothly
          chars = Math.max(3, Math.ceil(slack / 15));
        } else if (slack > 80) {
          chars = 3;   // Plenty of buffer, reveal at normal speed
        } else if (slack > 40) {
          chars = 2;   // Buffer shrinking, ease off
        } else if (slack > 15) {
          chars = 1;   // Running low, drip-feed to avoid catching up
        } else {
          chars = 0;   // Nearly caught up, wait for more SSE data
        }

        if (chars > 0 && revealedLen < targetText.length) {
          revealedLen = Math.min(revealedLen + chars, targetText.length);
          onChunk(targetText.substring(0, revealedLen));
        }

        if (streamDone && revealedLen >= targetText.length) {
          clearInterval(tick);
          resolve();
        }
      }, TICK_MS);
    });

    // ── SSE read loop ─────────────────────────────────────────────────────────
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;

      buffer += decoder.decode(value, { stream: true });

      // SSE format: "data: {JSON}\nid:N\n\n" — blank line separates events
      const parts = buffer.split('\n');
      buffer = parts.pop() ?? '';

      let jsonAccumulator = '';
      let inDataBlock = false;

      for (const line of parts) {
        if (line.startsWith('data: ')) {
          inDataBlock = true;
          jsonAccumulator = line.slice(6);
        } else if (inDataBlock && line.trim() === '') {
          // Blank line = end of this SSE event
          inDataBlock = false;
          if (jsonAccumulator.trim()) {
            try {
              const parsed = JSON.parse(jsonAccumulator);
              const messages: Array<{ text?: string }> = parsed?.messages ?? [];
              // Always take the LAST message — it is the assistant response.
              if (messages.length > 0) {
                const assistantText = (messages[messages.length - 1].text ?? '').trim();
                if (assistantText && assistantText.length > targetText.length) {
                  targetText = assistantText;
                }
              }
            } catch { /* skip malformed JSON */ }
          }
          jsonAccumulator = '';
        } else if (inDataBlock) {
          if (line.startsWith('id:') || line.startsWith('event:') || line.startsWith('retry:')) {
            continue;
          }
          jsonAccumulator += line;
        }
      }
    }

    // SSE stream finished — let the reveal timer catch up to the final text
    streamDone = true;
    await revealDone;

    if (!targetText) throw new Error('No assistant text received from stream');
    return targetText;
  }

  static generatePersonalizedFallbackText(
    displayName: string,
    jobTitle: string,
    organizationName: string
  ): string {
    const firstName = displayName.split(' ')[0];
    const rolePhrase = jobTitle
      ? ` As a ${jobTitle}, Copilot can help you draft documents, summarise meetings, and analyse data faster than ever.`
      : ' Copilot can help you draft documents, summarise meetings, and analyse data faster than ever.';
    return (
      `Hi **${firstName}**!${rolePhrase} ` +
      `Join the Copilot Engagement Program at ${organizationName || 'your organisation'} ` +
      `to track your progress, earn points, and climb the leaderboard — your AI journey starts now!`
    );
  }
}

