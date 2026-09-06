import { NextResponse } from "next/server";
import { validateMacroForm } from "@/lib/formValidation";
import { buildSpecification } from "@/lib/specBuilder";
import { buildPrompt } from "@/lib/promptBuilder";
import { AiProviderError, callAiProvider, isAiConfigured } from "@/lib/aiProvider";
import { parseAiResponse } from "@/lib/responseParser";
import { MacroFormData } from "@/lib/types";

// Default Node.js runtime (not edge) -- required for the OpenAI SDK.
export const runtime = "nodejs";

const MAX_BODY_BYTES = 200_000;

export async function POST(request: Request) {
  if (!isAiConfigured()) {
    return NextResponse.json(
      {
        ok: false,
        error: "AI generation is not configured on this server. Set OPENAI_API_KEY to enable it.",
        code: "missing-api-key",
      },
      { status: 503 }
    );
  }

  let bodyText: string;
  try {
    bodyText = await request.text();
  } catch {
    return NextResponse.json({ ok: false, error: "Could not read the request body." }, { status: 400 });
  }

  if (bodyText.length > MAX_BODY_BYTES) {
    return NextResponse.json({ ok: false, error: "Request body is too large." }, { status: 413 });
  }

  let form: MacroFormData;
  try {
    const parsed = JSON.parse(bodyText);
    form = parsed.form as MacroFormData;
  } catch {
    return NextResponse.json({ ok: false, error: "Request body was not valid JSON." }, { status: 400 });
  }

  if (!form) {
    return NextResponse.json({ ok: false, error: "Missing form data." }, { status: 400 });
  }

  const validation = validateMacroForm(form);
  if (!validation.valid) {
    return NextResponse.json(
      { ok: false, error: "The form has validation errors.", fieldErrors: validation.errors },
      { status: 400 }
    );
  }

  const spec = buildSpecification(form);
  const { system, user } = buildPrompt(spec);

  try {
    const raw = await callAiProvider(system, user);
    const parsedResponse = parseAiResponse(raw);

    if (!parsedResponse.ok || !parsedResponse.result) {
      return NextResponse.json(
        { ok: false, error: parsedResponse.error ?? "Could not parse the AI response.", specification: spec },
        { status: 502 }
      );
    }

    return NextResponse.json({ ok: true, specification: spec, result: parsedResponse.result });
  } catch (err) {
    if (err instanceof AiProviderError) {
      const status = err.code === "rate-limit" ? 429 : err.code === "timeout" ? 504 : 502;
      return NextResponse.json({ ok: false, error: err.message, code: err.code, specification: spec }, { status });
    }
    return NextResponse.json(
      { ok: false, error: "An unexpected server error occurred while generating the macro.", specification: spec },
      { status: 500 }
    );
  }
}
