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
export const PERSONALIZED_WELCOME_PROMPT = `Write a short, friendly 2–3 sentence welcome for an employee joining the "Copilot Engagement Program".

Rules:
- Address the user by their first name
- Reference their job role and/or department to make it feel tailored
- Mention that colleagues are already participating (use a natural phrase like "dozens of colleagues" or a similar plausible number)
- Tone: warm, encouraging, concise — never corporate jargon
- Length: 40–60 words maximum
- Output ONLY the welcome text — no titles, no quotes, no meta-commentary
- You MAY use **double asterisks** around ONE short key phrase for emphasis
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

