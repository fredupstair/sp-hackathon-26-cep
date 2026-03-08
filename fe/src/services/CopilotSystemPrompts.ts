// ─── Quick prompt suggestions ─────────────────────────────────────────────────

export interface IQuickPrompt {
  key: string;
  label: string;
  description: string;
  instruction: string;
}

export const QUICK_PROMPTS: IQuickPrompt[] = [
  {
    key: 'motivational',
    label: '🌟 Motivational',
    description: 'Inspire & energise your team',
    instruction:
      'Write an inspiring, motivational welcome text that energises employees and positions AI as a career growth opportunity.',
  },
  {
    key: 'data-driven',
    label: '📊 Data-driven',
    description: 'Facts, metrics & leaderboard wins',
    instruction:
      'Write a factual, results-focused welcome text that highlights measurable benefits — tracking usage, earning points, and leaderboard visibility.',
  },
  {
    key: 'innovation',
    label: '🚀 Innovation',
    description: 'Digital transformation & future of work',
    instruction:
      'Write a forward-looking welcome text that frames the program as part of the digital transformation journey and innovation culture.',
  },
  {
    key: 'community',
    label: '🤝 Community',
    description: 'Team spirit & learning together',
    instruction:
      'Write a warm, community-focused welcome text that emphasises team spirit, friendly competition, and learning together.',
  },
  {
    key: 'punchy',
    label: '⚡ Short & punchy',
    description: 'Concise, max 60 words',
    instruction:
      'Write a very concise (max 60 words) welcome text with a clear, energetic call to action — straight to the point.',
  },
  {
    key: 'casual',
    label: '😊 Casual & friendly',
    description: 'Relaxed, human, no corporate jargon',
    instruction:
      'Write a casual, conversational welcome text that feels human and approachable — as if written by a friendly colleague, with no corporate jargon.',
  },
];

// ─── System prompt ────────────────────────────────────────────────────────────

/**
 * System prompt for the Copilot Chat conversation used to generate
 * the opt-in wizard welcome text.
 */
/**
 * Short personalised welcome prompt — used at runtime for each end user.
 * Produces a 2–3 sentence message tailored to the individual's name, role, and department.
 */
export const PERSONALIZED_WELCOME_PROMPT = `The user is asking how Microsoft 365 Copilot can help them in their daily work. Answer with a short, practical reply tailored to their job role.

Rules:
- Start with a warm greeting using their first name — always wrap the first name in **bold** (e.g. **Giulia**)
- Give exactly 3 concrete examples of how Copilot can help them, based on their specific role and/or department (e.g. drafting emails, summarising meetings, analysing data, creating presentations)
- Each example should be a brief phrase, naturally woven into 2–3 flowing sentences — NOT a bulleted list
- Close with one short sentence inviting the user to join the Copilot Engagement Program to track their progress, earn points, and climb the leaderboard
- Tone: friendly, practical, encouraging — no corporate jargon
- Length: 50–80 words maximum
- Output ONLY the response text — no titles, no quotes, no meta-commentary
- Bold ONLY the user's first name — nothing else
- No citations, no hyperlinks, no footnotes`;

export const COPILOT_SYSTEM_PROMPT = `You are a professional copywriter helping organisations encourage employee adoption of Microsoft 365 Copilot.

Your task is to write a concise, engaging welcome text for an internal initiative called the "Copilot Engagement Program".

Guidelines:
- Length: 2–3 short paragraphs, 100–150 words total
- Tone: enthusiastic but professional, appropriate for internal corporate communication
- Clearly convey: what the programme is, why it matters, and what participants gain (points, levels, leaderboard, recognition)
- Always render the organisation name in **bold** (e.g. **Contoso**) every time it appears in the text
- Do NOT mention specific competing AI products or tools
- Do NOT add extra headings, bullet lists, or markdown — plain prose only
- You MAY use **double asterisks** around key words or short phrases to emphasise them (e.g. **earn points**, **leaderboard**) — this is the ONLY formatting allowed
- Do NOT include citations, footnotes, reference numbers, hyperlinks, or URLs of any kind (e.g. [1](...), [source](...)) — these are strictly forbidden
- Write the text directly, without meta-commentary like "Here is the text:"`;

