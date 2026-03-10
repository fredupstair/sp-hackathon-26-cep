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
    conversationId?: string,
    staticPrefix?: string,
  ): Promise<{ text: string; fromFallback: boolean }> {
    try {
      const convId = conversationId ?? await this.createConversation();
      const firstName = displayName.split(' ')[0];
      const roleContext = [jobTitle, department].filter(Boolean).join(', ');
      const userContext =
        `User: ${firstName}${roleContext ? ` (${roleContext})` : ''}. ` +
        `Organisation: "${organizationName || 'our company'}". Write the bullet points now.`;
      const fullPrompt = `${PERSONALIZED_WELCOME_PROMPT}\n\n${userContext}`;

      const text = await this._sendStreamMessage(convId, fullPrompt, onChunk, staticPrefix);
      return { text, fromFallback: false };
    } catch {
      // Streaming failed — fall back to non-streaming then static fallback
      try {
        const result = await this.generatePersonalizedText(displayName, jobTitle, department, organizationName);
        const combined = staticPrefix ? `${staticPrefix}\n${result.text}` : result.text;
        onChunk(combined);
        return { text: combined, fromFallback: result.fromFallback };
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
   *
   * When staticPrefix is provided it is typed out immediately (before the
   * HTTP response even arrives), so the user sees instant feedback. The
   * Copilot bullet-point response is then appended seamlessly.
   */
  private async _sendStreamMessage(
    conversationId: string,
    text: string,
    onChunk: (accumulated: string) => void,
    staticPrefix?: string,
  ): Promise<string> {
    const body = {
      message: { text },
      locationHint: { timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone },
      contextualResources: {
        webContext: { isWebEnabled: false },
      },
    };

    // ── State shared between reveal timer and SSE loop ────────────────────────
    let targetText  = staticPrefix ?? '';  // grows: staticPrefix + '\n' + copilot chunks
    let copilotText = '';                  // tracks just the Copilot contribution
    let revealedLen = 0;
    let streamDone  = false;
    const staticPrefixLen = staticPrefix?.length ?? 0;
    const TICK_MS = 40;   // ms between each reveal step
    // Prefix phase: 1 char / 40ms ≈ 25 chars/sec — comfortable typewriter feel
    // Copilot phase: 5 chars / 40ms ≈ 125 chars/sec — fast reveal of bullets

    // ── Reveal timer starts IMMEDIATELY so the static prefix types right away ─
    // setInterval fires between every async await, so it runs while the HTTP
    // request is in-flight — giving the impression of instant response.
    const revealDone = new Promise<void>((resolve) => {
      const tick = setInterval(() => {
        const slack = targetText.length - revealedLen;
        let chars: number;

        if (revealedLen < staticPrefixLen) {
          chars = 1;   // Slow typewriter for the static prefix
        } else if (streamDone) {
          chars = Math.max(3, Math.ceil(slack / 10));  // Flush bullets quickly
        } else if (slack > 80) {
          chars = 5;
        } else if (slack > 40) {
          chars = 3;
        } else if (slack > 15) {
          chars = 2;
        } else {
          chars = 0;   // Waiting for more SSE data
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

    // ── HTTP fetch + SSE loop run in parallel with the reveal timer ───────────
    // JS event loop: setInterval fires during every 'await', so the timer
    // types the static prefix while we wait for the API response.
    try {
      // Use responseType 'raw' to get the native Response for SSE parsing
      const rawResponse: Response = await (this._graphClient
        .api(`/copilot/conversations/${conversationId}/chatOverStream`)
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        .version('beta') as any)
        .responseType('raw')
        .post(body);

      if (!rawResponse.body) throw new Error('No response body from chatOverStream');

      const reader  = rawResponse.body.getReader();
      const decoder = new TextDecoder();
      let buffer = '';
      const sep = staticPrefix ? '\n' : '';

      // ── SSE read loop ───────────────────────────────────────────────────────
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
            inDataBlock = false;
            if (jsonAccumulator.trim()) {
              try {
                const parsed = JSON.parse(jsonAccumulator);
                const messages: Array<{ text?: string }> = parsed?.messages ?? [];
                if (messages.length > 0) {
                  const chunk = (messages[messages.length - 1].text ?? '').trim();
                  if (chunk && chunk.length > copilotText.length) {
                    copilotText = chunk;
                    targetText = `${staticPrefix ?? ''}${sep}${copilotText}`;
                  }
                }
              } catch { /* skip malformed JSON */ }
            }
            jsonAccumulator = '';
          } else if (inDataBlock) {
            if (line.startsWith('id:') || line.startsWith('event:') || line.startsWith('retry:')) continue;
            jsonAccumulator += line;
          }
        }
      }
    } catch { /* swallow — timer will flush whatever targetText we have so far */ }

    // Let the reveal timer catch up to the final targetText
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

