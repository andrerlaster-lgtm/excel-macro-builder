import { afterEach, describe, expect, it } from "vitest";
import { isAiConfigured } from "./aiProvider";

const ORIGINAL_KEY = process.env.OPENAI_API_KEY;

afterEach(() => {
  if (ORIGINAL_KEY === undefined) {
    delete process.env.OPENAI_API_KEY;
  } else {
    process.env.OPENAI_API_KEY = ORIGINAL_KEY;
  }
});

describe("isAiConfigured", () => {
  it("is false when OPENAI_API_KEY is unset", () => {
    delete process.env.OPENAI_API_KEY;
    expect(isAiConfigured()).toBe(false);
  });

  it("is false when OPENAI_API_KEY is blank", () => {
    process.env.OPENAI_API_KEY = "   ";
    expect(isAiConfigured()).toBe(false);
  });

  it("is true when OPENAI_API_KEY is set", () => {
    process.env.OPENAI_API_KEY = "sk-test-not-real";
    expect(isAiConfigured()).toBe(true);
  });
});
