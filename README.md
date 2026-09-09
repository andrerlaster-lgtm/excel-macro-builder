# Excel Macro Builder

A single-user, local-first web app that turns a plain-language Excel automation
problem — plus structured details about your workbook — into a **reviewable
VBA macro draft**. It uses OpenAI's API server-side to generate the code.

**This app never opens, uploads, edits, or executes any Excel workbook.** It
only produces text (VBA source + explanatory notes) for you to read and paste
into Excel yourself.

## What it does

1. **Describe the task** — the problem, the desired result, and how you'll
   know it worked.
2. **Map the Excel data** — source/destination workbook names, worksheets,
   ranges/tables, headers, match fields, filters, and transformation rules.
   Anything that doesn't apply to your task can be explicitly marked
   **Not applicable** instead of forcing you to invent details.
3. **Review and generate** — see the exact structured specification that will
   be sent for generation, resolve anything left ambiguous, preview a
   before/after example (below), then generate.

The generated result includes: the VBA code (editable), a plain-language
summary, assumptions the model made, open questions that could change the
result, expected inputs/outputs, install/run instructions, a test plan for a
**copy** of your workbook, safety cautions, and platform limitations.

## Before/after preview

Step 2 has an optional **Preview operation** picker (copy, filter, remove
duplicates, or group-and-total). When you choose one, step 3 shows a
**Before** and **After** table computed from a small editable sample grid,
seeded from the column headers and example rows you typed in step 2. You can
edit any cell, add or remove rows, or reseed from step 2, and the result
recomputes instantly.

- **No API key required.** The whole thing is a pure, deterministic function
  (`src/lib/previewSimulator.ts`) running in your browser. Nothing is sent
  anywhere, and your sample rows are deliberately excluded from the
  specification that goes to the AI provider — only the operation shape
  (kind, columns, aggregate function, filter operator) travels.
- **It does not execute the generated VBA.** This is the important part. The
  preview simulates *the operation you described with the picker*, not the
  code the model wrote. It cannot prove the macro behaves the same way, and
  the on-screen amber caution says so. The free-text rule fields on step 2
  remain the source of truth for generation, and the real macro can do
  something the preview does not show. Still read the code and test it on a
  copy of your workbook.
- Stages apply in a fixed order: filter → deduplicate → aggregate, with
  aggregate/deduplicate output sorted by key ascending. Deduplicate always
  keeps the *first* row per key, which may differ from your free-text
  "Duplicate handling" rule; the preview says so in its step narrative.
- Non-numeric values in an aggregated column are **skipped and reported**,
  never silently coerced to zero. Currency and thousands separators
  (`$1,234.50`) parse correctly. Simulation is capped at 500 rows and 50
  columns, and truncation is noted.

## Setup

```bash
npm install
cp .env.example .env.local
# edit .env.local and set OPENAI_API_KEY
npm run dev
```

### API key and cost

- Set `OPENAI_API_KEY` in `.env.local` (never commit a real key — `.env*` is
  gitignored except `.env.example`).
- Get a key at <https://platform.openai.com/api-keys>.
- **A ChatGPT Plus/Team/Pro subscription does NOT include API credits.** API
  usage is billed separately, per token, on your OpenAI platform billing
  account. Set a spending limit at
  <https://platform.openai.com/account/limits> before heavy use.
- Every click of "Generate VBA" sends your typed form content (never anything
  else — no vault files, no env vars, no hidden system data) to OpenAI. Don't
  enter confidential information unless you're authorized to send it to a
  third-party API.

### Running without an API key

The app is fully usable with no key configured: the guided form, validation,
and the structured specification (copy/download as JSON) all work. Only the
"Generate VBA" action is disabled, with an honest on-screen message —
**the app never pretends code was generated when it wasn't.**

## Architecture

Everything is a standalone Next.js App Router project (TypeScript, no
backend database). Layers are kept separable and unit-tested:

- `src/lib/types.ts` — form/data model, specification shape, saved-project shape.
- `src/lib/formValidation.ts`, `src/lib/vbaValidation.ts` — field and VBA
  identifier validation.
- `src/lib/specBuilder.ts` — deterministically builds a structured
  specification object from the form (no free-text paragraph).
- `src/lib/promptBuilder.ts` — turns the specification into the system/user
  prompt text sent to the model, including the VBA quality rules below.
- `src/lib/aiProvider.ts` — the **only** module that touches the OpenAI SDK
  and the model name, so the provider can be swapped later without touching
  anything else. Server-side only.
- `src/lib/responseParser.ts` — defensively parses the model's JSON response;
  never evals/executes anything; fails honestly on malformed output.
- `src/lib/safetyScanner.ts` — a pure, unit-tested static pattern scan over
  generated VBA text (see Safety below).
- `src/lib/previewSimulator.ts` — pure, deterministic before/after simulation
  of the structured preview operation. No DOM, no network, no VBA execution.
- `src/lib/filename.ts` — derives a safe `.bas` filename.
- `src/lib/storage.ts` — versioned browser localStorage for saved projects,
  with a migration function for future schema changes.
- `src/app/api/generate/route.ts` — the only server route that calls OpenAI
  (Node.js runtime, not edge — required for the SDK). Re-validates the form
  server-side before calling the provider.
- `src/components/*` — the guided three-step UI, before/after preview panel,
  result panel, and project sidebar.

## Local data / storage limitations

- Draft projects and generated results are stored in this browser's
  `localStorage` only — no sign-in, no cloud sync, no backend database.
  Clearing browser data or switching browsers/devices loses everything.
- There is **no permanent delete** in this phase. Projects can be archived
  and restored, but not erased from the UI. To fully remove all data, use
  your browser's own site-data controls (e.g. in Chrome: Settings → Privacy
  and security → Site settings → find this site → Clear data; in Safari:
  Settings → Privacy → Manage Website Data).
- The storage schema has an explicit integer version and a migration
  function (`migrateProject` in `src/lib/storage.ts`), exercised by a test,
  so future shape changes won't crash on old saved data.

## Macro installation (what you do with the output)

1. Open Excel and press **Alt+F11** (Windows) or use the Developer tab
   (Mac) to open the VBA editor.
2. Insert a standard module (Insert → Module).
3. Paste the generated code in.
4. Save the workbook as a macro-enabled file: **File → Save As → Excel
   Macro-Enabled Workbook (.xlsm)**. A macro cannot be saved in a plain
   `.xlsx` file.
5. Run it from the VBA editor (F5), an assigned button, or the configured
   trigger — after reading it and testing on a **copy** of your workbook.

## Safety boundaries

- **Review before running** — this is shown throughout the app. Generated
  code can be wrong, incomplete, or destructive even if it looks plausible.
  Nothing produced here is verified safe.
- **Static warning pass** (`src/lib/safetyScanner.ts`) flags patterns like
  `Kill`, `RmDir`, `FileSystemObject.DeleteFile`, `Shell`, bulk row/column
  deletion, `.Save`/`.SaveAs`, Outlook/`SendMail` automation, and external
  network calls (`WinHttp`, `XMLHTTP`, `URLDownloadToFile`). **This is a
  review aid, not a safety proof — no warning does not mean the code is
  safe.**
- The app **never verifies** that any workbook path, sheet, range, or table
  you typed actually exists. All of that is generation context only.
- Windows and Mac Excel differ (file dialogs, paths, ActiveX, Outlook
  automation, some external integrations). The selected platform flows into
  the prompt, and the model is asked to surface platform-specific
  limitations in its output — but this cannot be guaranteed for every
  response from a live LLM call.
- VBA quality rules given to the model (and enforced by prompt instructions,
  not a compiler): `Option Explicit`; fully qualified workbook/worksheet/
  range/table references (no `Select`/`Selection`/`Activate`/`ActiveSheet`);
  a clear configuration section; descriptive names and concise comments; an
  error handler with a cleanup path; restoring changed `Application` state
  (Calculation, EnableEvents, DisplayAlerts, ScreenUpdating); header-based
  column lookup when headers were supplied; explicit handling of missing
  workbooks/sheets/ranges/tables/headers. These are instructions to the
  model, not a guarantee — always read the generated code.
- Event-driven triggers (on workbook open / on sheet change) get an extra
  on-screen caution because they run automatically without you clicking
  anything.

## Testing

```bash
npm run lint
npm test
npm run build
```

Vitest covers: form validation, VBA identifier validation, structured
specification generation (including that every entered field is carried
through unmodified and "Not applicable" is never invented), AI response
parsing (malformed/partial JSON handled honestly), the static
dangerous-pattern scanner (one fixture per pattern category), safe `.bas`
filename generation, localStorage migration of a legacy record, and the
no-API-key fallback (`isAiConfigured`).

## Known limitations / next steps

- No file upload in this phase — everything is typed text, by design.
- No automated end-to-end test of a real OpenAI call (would cost money and
  isn't deterministic); the parser and route are tested with fixtures
  instead.
- Recommended next step: add a lightweight, non-billed "is AI configured?"
  status check the UI can call on load, instead of only discovering it on
  the first Generate click.
