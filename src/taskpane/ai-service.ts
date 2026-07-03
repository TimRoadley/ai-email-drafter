/* global fetch */

import { buildPrompt, buildClarityPrompt } from "./prompt-template";

export const DEFAULT_BASE_URL = "https://api.example.com";

export interface AiStatus {
  online: boolean;
  models: string[];
  error?: string;
}

function buildHeaders(apiKey?: string): Record<string, string> {
  const headers: Record<string, string> = {
    "Content-Type": "application/json",
  };
  if (apiKey && apiKey.trim()) {
    headers["Authorization"] = `Bearer ${apiKey.trim()}`;
  }
  return headers;
}

export async function checkAiStatus(
  baseUrl: string = DEFAULT_BASE_URL,
  apiKey?: string,
): Promise<AiStatus> {
  try {
    const response = await fetch(`${baseUrl}/models`, {
      headers: buildHeaders(apiKey),
    });
    if (!response.ok) {
      throw new Error(`Server returned ${response.status}`);
    }
    const data = await response.json();
    const models = data.data?.map((m: any) => m.id) || [];
    return { online: true, models };
  } catch (error: any) {
    return { online: false, models: [], error: error.message };
  }
}

export async function generateReply(
  baseUrl: string,
  model: string,
  instructions: string,
  emailBody: string,
  thinkingEnabled: boolean = true,
  customPromptTemplate?: string,
  signal?: AbortSignal,
  apiKey?: string,
): Promise<string> {
  const prompt = buildPrompt(
    instructions,
    emailBody,
    customPromptTemplate,
    thinkingEnabled,
  );

  const fetchOptions: RequestInit = {
    method: "POST",
    headers: buildHeaders(apiKey),
    body: JSON.stringify({
      model: model,
      messages: [{ role: "user", content: prompt }],
      temperature: 0.7,
    }),
  };

  if (signal) {
    fetchOptions.signal = signal;
  }

  const response = await fetch(`${baseUrl}/chat/completions`, fetchOptions);

  if (!response.ok) {
    let detail = "";
    try {
      const errBody = await response.json();
      detail =
        errBody?.error?.message || errBody?.message || JSON.stringify(errBody);
    } catch {
      detail = await response.text().catch(() => "");
    }
    throw new Error(
      `AI request failed: ${response.status} ${response.statusText}${detail ? ` — ${detail}` : ""}`,
    );
  }

  const data = await response.json();
  return data.choices[0].message.content;
}

export async function rewriteTextForClarity(
  baseUrl: string,
  model: string,
  text: string,
  apiKey?: string,
): Promise<string> {
  const prompt = buildClarityPrompt(text);

  const response = await fetch(`${baseUrl}/chat/completions`, {
    method: "POST",
    headers: buildHeaders(apiKey),
    body: JSON.stringify({
      model: model,
      messages: [{ role: "user", content: prompt }],
      temperature: 0.4,
    }),
  });

  if (!response.ok) {
    let detail = "";
    try {
      const errBody = await response.json();
      detail =
        errBody?.error?.message || errBody?.message || JSON.stringify(errBody);
    } catch {
      detail = await response.text().catch(() => "");
    }
    throw new Error(
      `AI request failed: ${response.status} ${response.statusText}${detail ? ` — ${detail}` : ""}`,
    );
  }

  const data = await response.json();
  return data.choices[0].message.content;
}
