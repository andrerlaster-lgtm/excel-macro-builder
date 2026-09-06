import { MacroSpecification } from "./types";

export const RESPONSE_JSON_SCHEMA_DESCRIPTION = `
Return ONLY a single JSON object (no markdown fences, no prose outside it) with exactly this shape:
{
  "vbaCode": string,            // complete VBA module source, ready to paste into a .bas module
  "summary": string,            // plain-language summary of what the macro does
  "assumptions": string[],      // assumptions made to fill gaps in the spec
  "openQuestions": string[],    // missing details that could materially change the result
  "inputsOutputs": string,      // expected inputs and outputs, in plain language
  "installInstructions": string, // how to install/run it, including saving as .xlsm when needed
  "testPlan": string[],         // steps to test using a COPY of the workbook
  "safetyCautions": string[],   // safety cautions and platform limitations
  "platformLimitations": string[] // platform-specific caveats (Windows vs Mac Excel)
}`.trim();

const QUALITY_RULES = `
VBA quality rules to follow unless the task genuinely requires otherwise:
- Start every module with "Option Explicit".
- Use fully qualified references to workbooks, worksheets, ranges, and tables. Do not rely on ActiveWorkbook, ActiveSheet, Selection, or the Select/Activate methods.
- Put all configurable values (workbook names, sheet names, table/range names, header row number, column headers) in one clearly marked configuration section near the top.
- Use descriptive variable and procedure names, and add concise comments only where logic is non-obvious.
- Wrap the main procedure in an error handler with a cleanup path (a single exit point that always runs).
- Save and restore any Application state you change (Calculation, EnableEvents, DisplayAlerts, ScreenUpdating) in the cleanup path, even on error.
- When column headers were supplied, look up columns by header text instead of hard-coded column letters/numbers.
- Explicitly check for and handle missing workbooks, worksheets, ranges, tables, or required headers, with a clear message instead of a raw runtime error.
`.trim();

/**
 * Builds the deterministic prompt sent to the AI provider from the
 * structured specification. This is pure text construction -- no model
 * call happens here.
 */
export function buildPrompt(spec: MacroSpecification): { system: string; user: string } {
  const system = [
    "You are an expert VBA developer who writes safe, explicit, review-friendly Excel macros.",
    "You NEVER execute code yourself. You only produce VBA source text and explanatory text for a human to review before running it themselves.",
    "You must surface any assumption or missing detail explicitly rather than silently guessing and hiding it.",
    QUALITY_RULES,
    RESPONSE_JSON_SCHEMA_DESCRIPTION,
  ].join("\n\n");

  const user = `Generate a VBA macro from this structured specification. Treat every field literally; "Not applicable" means the user intentionally left that concern out of scope.

${JSON.stringify(spec, null, 2)}

Target Excel platform: ${spec.platform}. Reflect platform-specific limitations (file dialogs, paths, ActiveX, Outlook automation, other external integrations differ between Windows and Mac Excel) in platformLimitations.

Trigger mode: ${spec.execution.trigger}. If this is "workbook-open" or "sheet-change" (event-driven), include an explicit extra caution in safetyCautions about event-driven macros running automatically and the risk of unexpected side effects, and explain where the event code must be placed (e.g. the ThisWorkbook or worksheet code module) versus a standard module.

Respond with the JSON object described in the system message only.`;

  return { system, user };
}
