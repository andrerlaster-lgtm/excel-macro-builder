import { describe, expect, it } from "vitest";
import { safeBasFilename } from "./filename";

describe("safeBasFilename", () => {
  it("appends .bas to a clean name", () => {
    expect(safeBasFilename("UpdateMonthlySummary")).toBe("UpdateMonthlySummary.bas");
  });

  it("replaces spaces and punctuation with underscores", () => {
    expect(safeBasFilename("Update Sales! (final)")).toBe("Update_Sales_final.bas");
  });

  it("falls back to macro.bas when nothing usable remains", () => {
    expect(safeBasFilename("!!!")).toBe("macro.bas");
    expect(safeBasFilename("")).toBe("macro.bas");
  });

  it("avoids Windows reserved device names", () => {
    expect(safeBasFilename("con")).toBe("con_macro.bas");
    expect(safeBasFilename("LPT1")).toBe("LPT1_macro.bas");
  });

  it("collapses repeated underscores", () => {
    expect(safeBasFilename("a___b")).toBe("a_b.bas");
  });

  it("truncates very long names", () => {
    const result = safeBasFilename("A".repeat(500));
    expect(result.length).toBeLessThanOrEqual(104);
    expect(result.endsWith(".bas")).toBe(true);
  });
});
