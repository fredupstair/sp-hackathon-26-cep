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

  static generatePersonalizedFallbackText(
    displayName: string,
    jobTitle: string,
    organizationName: string
  ): string {
    const firstName = displayName.split(' ')[0];
    const rolePhrase = jobTitle ? ` As a ${jobTitle}, you` : ' You';
    return (
      `Hi ${firstName}! 👋${rolePhrase} can use Copilot every day to work smarter and save hours every week. ` +
      `Join dozens of colleagues at **${organizationName || 'your organisation'}** ` +
      `already earning points and climbing the leaderboard — your AI journey starts now!`
    );
  }
}

