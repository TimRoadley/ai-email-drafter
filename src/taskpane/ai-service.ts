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

/**
 * Extracts the assistant's reply text from a chat/completions response body.
 *
 * Some OpenAI-compatible servers/models don't return a plain string for
 * `choices[0].message.content` — reasoning-only models may leave it `null`
 * (with the real text elsewhere), and some servers return an array of
 * content parts (e.g. `[{ type: "text", text: "..." }]`) instead of a
 * string. Passing a non-string straight into Office.js APIs like
 * `setSelectedDataAsync` throws a confusing `Sys.ArgumentTypeException`, so
 * normalize here and fail with a clear error instead.
 */
function extractMessageContent(data: any): string {
  const content = data?.choices?.[0]?.message?.content;

  if (typeof content === "string" && content.length > 0) {
    return content;
  }

  if (Array.isArray(content)) {
    const text = content
      .filter((part: any) => typeof part?.text === "string")
      .map((part: any) => part.text)
      .join("");
    if (text) {
      return text;
    }
  }

  throw new Error(
    "The model returned no usable text content. Try a different model or check the endpoint's response format.",
  );
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
    const models = (data.data?.map((m: any) => m.id) || []).sort();
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
  return extractMessageContent(data);
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
  return extractMessageContent(data);
}
