import { describe, expect, it } from "vitest";
import { validateVbaMacroName } from "./vbaValidation";

describe("validateVbaMacroName", () => {
  it("accepts a simple valid identifier", () => {
    expect(validateVbaMacroName("UpdateMonthlySummary").valid).toBe(true);
  });

  it("accepts underscores and digits after the first letter", () => {
    expect(validateVbaMacroName("Sync_Sales_2024").valid).toBe(true);
  });

  it("rejects an empty name", () => {
    const result = validateVbaMacroName("");
    expect(result.valid).toBe(false);
    expect(result.errors.length).toBeGreaterThan(0);
  });

  it("rejects a name with spaces", () => {
    expect(validateVbaMacroName("Update Sales").valid).toBe(false);
  });

  it("rejects a name with punctuation", () => {
    expect(validateVbaMacroName("Update-Sales!").valid).toBe(false);
  });

  it("rejects a name starting with a digit", () => {
    expect(validateVbaMacroName("2024Sync").valid).toBe(false);
  });

  it("rejects a reserved VBA keyword", () => {
    expect(validateVbaMacroName("Function").valid).toBe(false);
    expect(validateVbaMacroName("If").valid).toBe(false);
    expect(validateVbaMacroName("dim").valid).toBe(false);
  });

  it("rejects a name that is too long", () => {
    const longName = "A".repeat(41);
    expect(validateVbaMacroName(longName).valid).toBe(false);
  });

  it("rejects leading/trailing whitespace", () => {
    expect(validateVbaMacroName(" UpdateSales").valid).toBe(false);
    expect(validateVbaMacroName("UpdateSales ").valid).toBe(false);
  });
});
