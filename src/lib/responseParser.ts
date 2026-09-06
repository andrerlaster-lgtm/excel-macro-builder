import { MacroGenerationResult } from "./types";

export interface ParseResult {
  ok: boolean;
  result: MacroGenerationResult | null;
  error: string | null;
}

function asStringArray(v: unknown): string[] {
  if (!Array.isArray(v)) return [];
  return v.filter((x): x is string => typeof x === "string" && x.trim().length > 0);
}

function asString(v: unknown): string {
  return typeof v === "string" ? v : "";
}

/**
 * Defensively parses AI provider output (expected to be a JSON object) into
 * a MacroGenerationResult. Treats the model output as untrusted text:
 * malformed JSON or a shape mismatch produces an honest failure, never a
 * thrown exception and never executed/evaluated code.
 */
export function parseAiResponse(raw: string): ParseResult {
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    return { ok: false, result: null, error: "The AI response was not valid JSON and could not be parsed." };
  }

  if (typeof parsed !== "object" || parsed === null) {
    return { ok: false, result: null, error: "The AI response JSON was not an object." };
  }

  const obj = parsed as Record<string, unknown>;
  const vbaCode = asString(obj.vbaCode);

  if (vbaCode.trim().length === 0) {
    return { ok: false, result: null, error: "The AI response did not include any VBA code." };
  }

  const result: MacroGenerationResult = {
    vbaCode,
    summary: asString(obj.summary),
    assumptions: asStringArray(obj.assumptions),
    openQuestions: asStringArray(obj.openQuestions),
    inputsOutputs: asString(obj.inputsOutputs),
    installInstructions: asString(obj.installInstructions),
    testPlan: asStringArray(obj.testPlan),
    safetyCautions: asStringArray(obj.safetyCautions),
    platformLimitations: asStringArray(obj.platformLimitations),
  };

  return { ok: true, result, error: null };
}
