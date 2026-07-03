export const DEFAULT_PROMPT_TEMPLATE: string = `You are an email assistant. Draft a professional email based on the instructions and context provided.

[OUTPUT FORMAT]
- Start directly with: Hi {{firstName}}, (Only if the recipient's name is known. Otherwise, start with the body text).
- Followed by a blank line.
- Write the email body using short paragraphs separated by blank lines.
- Stop immediately after the final sentence. 

[STRICT RESTRAINTS]
- Output ONLY the email text. 
- Never include a subject line, preamble, commentary, sign-off, or signature.
- Cut all fluff. Be brief, direct, and professional.

[INPUT CONTEXT]
Instructions: {{specificInstructions}}
Reference Email/Trail (if any): {{emailBody}}

[DRAFT]
`;

export function buildPrompt(
  specificInstructions: string,
  emailBody: string,
  customPromptTemplate?: string,
  thinkingEnabled: boolean = true,
): string {
  const template = customPromptTemplate || DEFAULT_PROMPT_TEMPLATE;

  let prompt = template
    .replace(/\{\{specificInstructions\}\}/g, specificInstructions)
    .replace(/\{\{emailBody\}\}/g, emailBody);

  if (!thinkingEnabled) {
    prompt += `\n\nEFFICIENCY DIRECTIVE — follow without exception:\n- Answer directly and concisely.\n- Do NOT explain your reasoning, chain of thought, or step-by-step thinking.\n- Do NOT use phrases like 'Let me think', 'First, I will', 'To arrive at this'.\n- Output ONLY the final result.`;
  }

  return prompt;
}

export function buildClarityPrompt(text: string): string {
  return `Rewrite the following text for maximum clarity. Remove ambiguity, fix awkward phrasing, and cut unnecessary words. Preserve the original meaning and approximate length. Output ONLY the rewritten text — no preamble, no explanation, no alternatives.

Text to rewrite:
${text}`;
}
