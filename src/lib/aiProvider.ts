// Single adapter module for the AI provider call. Keep the exact model name
// and provider SDK usage isolated here so the provider can be swapped later
// without touching prompt building, parsing, or the API route's control
// flow. This module must only ever run server-side.

import OpenAI from "openai";

// Keep the model name in exactly one place.
export const OPENAI_MODEL = "gpt-4o-mini";

export class AiProviderError extends Error {
  code: "missing-api-key" | "timeout" | "rate-limit" | "malformed-response" | "upstream-error";

  constructor(code: AiProviderError["code"], message: string) {
    super(message);
    this.code = code;
    this.name = "AiProviderError";
  }
}

export function isAiConfigured(): boolean {
  return Boolean(process.env.OPENAI_API_KEY && process.env.OPENAI_API_KEY.trim().length > 0);
}

/**
 * Calls the OpenAI Chat Completions API with JSON-mode enabled and returns
 * the raw text content. Throws AiProviderError with an actionable,
 * non-leaky message on failure. Callers are responsible for parsing/
 * validating the returned JSON text against the expected shape.
 */
export async function callAiProvider(system: string, user: string): Promise<string> {
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    throw new AiProviderError("missing-api-key", "AI generation is not configured on this server.");
  }

  const client = new OpenAI({ apiKey, timeout: 60_000 });

  try {
    const completion = await client.chat.completions.create({
      model: OPENAI_MODEL,
      response_format: { type: "json_object" },
      temperature: 0.2,
      max_tokens: 4000,
      messages: [
        { role: "system", content: system },
        { role: "user", content: user },
      ],
    });

    const content = completion.choices[0]?.message?.content;
    if (!content) {
      throw new AiProviderError("malformed-response", "The AI provider returned an empty response.");
    }
    return content;
  } catch (err) {
    if (err instanceof AiProviderError) throw err;

    const e = err as { status?: number; code?: string; name?: string };
    if (e?.name === "APIConnectionTimeoutError" || e?.code === "ETIMEDOUT") {
      throw new AiProviderError("timeout", "The AI provider timed out. Please try again.");
    }
    if (e?.status === 429) {
      throw new AiProviderError("rate-limit", "The AI provider rate-limited this request. Please wait and try again.");
    }
    if (e?.status && e.status >= 500) {
      throw new AiProviderError("upstream-error", "The AI provider had a server error. Please try again shortly.");
    }
    if (e?.status === 401 || e?.status === 403) {
      throw new AiProviderError("missing-api-key", "The configured OpenAI API key was rejected. Check the server's OPENAI_API_KEY.");
    }

    throw new AiProviderError("upstream-error", "The AI provider request failed unexpectedly.");
  }
}
