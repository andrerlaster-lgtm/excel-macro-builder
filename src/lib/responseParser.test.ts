import { describe, expect, it } from "vitest";
import { parseAiResponse } from "./responseParser";

const validPayload = {
  vbaCode: "Option Explicit\nSub Test()\nEnd Sub",
  summary: "Copies rows from one sheet to another.",
  assumptions: ["Header row is row 1."],
  openQuestions: ["What should happen to blank rows?"],
  inputsOutputs: "Input: Raw Data sheet. Output: Monthly Summary sheet.",
  installInstructions: "Paste into a standard module and save as .xlsm.",
  testPlan: ["Run on a copy of Sample_Sales.xlsx first."],
  safetyCautions: ["Test on a backup copy before running on real data."],
  platformLimitations: ["File dialogs differ between Windows and Mac Excel."],
};

describe("parseAiResponse", () => {
  it("parses a well-formed JSON payload", () => {
    const result = parseAiResponse(JSON.stringify(validPayload));
    expect(result.ok).toBe(true);
    expect(result.result?.vbaCode).toContain("Option Explicit");
    expect(result.result?.assumptions).toEqual(["Header row is row 1."]);
  });

  it("attributes every parsed result to the AI generator", () => {
    // Attribution is set here so template output can never be mistaken for
    // model output, or the other way round, once a project is reopened.
    const result = parseAiResponse(JSON.stringify(validPayload));
    expect(result.result?.generator).toBe("ai");
  });

  it("ignores a generator claimed inside the model's own JSON", () => {
    const result = parseAiResponse(JSON.stringify({ ...validPayload, generator: "template" }));
    expect(result.result?.generator).toBe("ai");
  });

  it("fails gracefully on malformed JSON", () => {
    const result = parseAiResponse("{ this is not json");
    expect(result.ok).toBe(false);
    expect(result.result).toBeNull();
    expect(result.error).toBeTruthy();
  });

  it("fails gracefully when JSON is not an object", () => {
    const result = parseAiResponse(JSON.stringify(["a", "b"]));
    expect(result.ok).toBe(false);
  });

  it("fails when vbaCode is missing or empty", () => {
    const { vbaCode: _vbaCode, ...rest } = validPayload;
    const result = parseAiResponse(JSON.stringify(rest));
    expect(result.ok).toBe(false);
    expect(result.error).toMatch(/VBA code/i);
  });

  it("defaults missing array/string fields instead of throwing", () => {
    const result = parseAiResponse(JSON.stringify({ vbaCode: "Sub A()\nEnd Sub" }));
    expect(result.ok).toBe(true);
    expect(result.result?.assumptions).toEqual([]);
    expect(result.result?.summary).toBe("");
  });

  it("filters out non-string entries from array fields", () => {
    const result = parseAiResponse(
      JSON.stringify({ ...validPayload, assumptions: ["ok", 5, null, "also ok"] })
    );
    expect(result.result?.assumptions).toEqual(["ok", "also ok"]);
  });
});
