/* global Office console localStorage */

const SETTINGS_KEY = "ai-reply-base-url";
const SIGNATURE_KEY = "ai-reply-signature";
const MODEL_KEY = "ai-reply-model";
const API_KEY = "ai-reply-api-key";

export function saveBaseUrl(url: string): void {
  try {
    Office.context.roamingSettings.set(SETTINGS_KEY, url);
    Office.context.roamingSettings.saveAsync(
      (asyncResult: Office.AsyncResult<void>) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Failed to save settings:", asyncResult.error.message);
        }
      },
    );
  } catch {
    // Fallback to localStorage if RoamingSettings isn't available
    localStorage.setItem(SETTINGS_KEY, url);
  }
}

export function saveSignature(value: string): void {
  try {
    Office.context.roamingSettings.set(SIGNATURE_KEY, value);
    Office.context.roamingSettings.saveAsync(
      (asyncResult: Office.AsyncResult<void>) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Failed to save signature:", asyncResult.error.message);
        }
      },
    );
  } catch {
    localStorage.setItem(SIGNATURE_KEY, value);
  }
}

export function getSignature(): string {
  try {
    const value = Office.context.roamingSettings.get(SIGNATURE_KEY);
    if (value != null) return String(value);
  } catch {
    // fallback
  }
  return localStorage.getItem(SIGNATURE_KEY) ?? "";
}

export function saveSelectedModel(model: string): void {
  try {
    Office.context.roamingSettings.set(MODEL_KEY, model);
    Office.context.roamingSettings.saveAsync(
      (asyncResult: Office.AsyncResult<void>) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Failed to save model:", asyncResult.error.message);
        }
      },
    );
  } catch {
    localStorage.setItem(MODEL_KEY, model);
  }
}

export function getSelectedModel(): string | null {
  try {
    const model = Office.context.roamingSettings.get(MODEL_KEY);
    if (model) return String(model);
  } catch {
    // fallback
  }
  return localStorage.getItem(MODEL_KEY);
}

export function getBaseUrl(): string | null {
  try {
    const url = Office.context.roamingSettings.get(SETTINGS_KEY);
    if (url) {
      return String(url);
    }
  } catch {
    // Fallback
  }
  return localStorage.getItem(SETTINGS_KEY);
}

const THINKING_KEY = "ai-reply-thinking";

export function saveThinkingEnabled(value: boolean): void {
  const str = String(value);
  try {
    Office.context.roamingSettings.set(THINKING_KEY, str);
    Office.context.roamingSettings.saveAsync(
      (asyncResult: Office.AsyncResult<void>) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error(
            "Failed to save thinking setting:",
            asyncResult.error.message,
          );
        }
      },
    );
  } catch {
    localStorage.setItem(THINKING_KEY, str);
  }
}

export function getThinkingEnabled(): boolean {
  try {
    const value = Office.context.roamingSettings.get(THINKING_KEY);
    if (value != null) return value === "true" || value === true;
  } catch {
    // fallback
  }
  const fallback = localStorage.getItem(THINKING_KEY);
  if (fallback != null) return fallback === "true";
  return false; // default: OFF for fast responses
}

const PROMPT_TEMPLATE_KEY = "ai-reply-prompt-template";

export function saveCustomPromptTemplate(template: string | null): void {
  try {
    Office.context.roamingSettings.set(PROMPT_TEMPLATE_KEY, template ?? "");
    Office.context.roamingSettings.saveAsync(
      (asyncResult: Office.AsyncResult<void>) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error(
            "Failed to save prompt template:",
            asyncResult.error.message,
          );
        }
      },
    );
  } catch {
    if (template != null) {
      localStorage.setItem(PROMPT_TEMPLATE_KEY, template);
    } else {
      localStorage.removeItem(PROMPT_TEMPLATE_KEY);
    }
  }
}

export function getCustomPromptTemplate(): string | null {
  try {
    const value = Office.context.roamingSettings.get(PROMPT_TEMPLATE_KEY);
    if (value != null) return String(value);
  } catch {
    // fallback
  }
  return localStorage.getItem(PROMPT_TEMPLATE_KEY);
}

export function saveApiKey(value: string): void {
  try {
    Office.context.roamingSettings.set(API_KEY, value);
    Office.context.roamingSettings.saveAsync(
      (asyncResult: Office.AsyncResult<void>) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Failed to save API key:", asyncResult.error.message);
        }
      },
    );
  } catch {
    localStorage.setItem(API_KEY, value);
  }
}

export function getApiKey(): string {
  try {
    const value = Office.context.roamingSettings.get(API_KEY);
    if (value != null) return String(value);
  } catch {
    // fallback
  }
  return localStorage.getItem(API_KEY) ?? "";
}
