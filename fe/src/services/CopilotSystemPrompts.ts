// ─── Quick prompt suggestions ─────────────────────────────────────────────────

export interface IQuickPrompt {
  key: string;
  label: string;
  instruction: string;
}

export const QUICK_PROMPTS: IQuickPrompt[] = [
  {
    key: 'motivational',
    label: '🌟 Motivational',
    instruction:
      'Write an inspiring, motivational welcome text that energises employees and positions AI as a career growth opportunity.',
  },
  {
    key: 'data-driven',
    label: '📊 Data-driven',
    instruction:
      'Write a factual, results-focused welcome text that highlights measurable benefits — tracking usage, earning points, and leaderboard visibility.',
  },
  {
    key: 'innovation',
    label: '🚀 Innovation',
    instruction:
      'Write a forward-looking welcome text that frames the program as part of the digital transformation journey and innovation culture.',
  },
  {
    key: 'community',
    label: '🤝 Community',
    instruction:
      'Write a warm, community-focused welcome text that emphasises team spirit, friendly competition, and learning together.',
  },
  {
    key: 'punchy',
    label: '⚡ Short & punchy',
    instruction:
      'Write a very concise (max 60 words) welcome text with a clear, energetic call to action — straight to the point.',
  },
];

// ─── System prompt ────────────────────────────────────────────────────────────

/**
 * System prompt for the Copilot Chat conversation used to generate
 * the opt-in wizard welcome text.
 */
export const COPILOT_SYSTEM_PROMPT = `You are a professional copywriter helping organisations encourage employee adoption of Microsoft 365 Copilot.

Your task is to write a concise, engaging welcome text for an internal initiative called the "Copilot Engagement Program".

Guidelines:
- Length: 2–3 short paragraphs, 100–150 words total
- Tone: enthusiastic but professional, appropriate for internal corporate communication
- Clearly convey: what the programme is, why it matters, and what participants gain (points, levels, leaderboard, recognition)
- Do NOT mention specific competing AI products or tools
- Do NOT add extra headings, bullet lists, or markdown — plain prose only
- You MAY use **double asterisks** around key words or short phrases to emphasise them (e.g. **earn points**, **leaderboard**) — this is the ONLY formatting allowed
- Do NOT include citations, footnotes, reference numbers, hyperlinks, or URLs of any kind (e.g. [1](...), [source](...)) — these are strictly forbidden
- Write the text directly, without meta-commentary like "Here is the text:"`;
