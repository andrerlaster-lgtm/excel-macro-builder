"use client";

import { MacroFormData, Maybe } from "@/lib/types";
import { buildSpecification, summarizeSpecification } from "@/lib/specBuilder";

interface UnresolvedField {
  label: string;
}

function isUnresolved(m: Maybe): boolean {
  return !m.notApplicable && m.value.trim().length === 0;
}

function collectUnresolved(form: MacroFormData): UnresolvedField[] {
  const checks: [Maybe, string][] = [
    [form.mapping.sourceExampleRows, "Example rows"],
    ...(form.mapping.sameWorkbook ? [] : ([[form.mapping.destinationWorkbook, "Destination workbook"]] as [Maybe, string][])),
    [form.mapping.matchField, "Match/lookup field"],
    [form.mapping.filters, "Filters"],
    [form.mapping.transformationRules, "Transformation/calculation rules"],
    [form.mapping.sortOrder, "Sort order"],
    [form.mapping.duplicateHandling, "Duplicate handling"],
    [form.mapping.blankOrErrorHandling, "Blank/error handling"],
    [form.run.constraints, "Optional constraints"],
    [form.run.mustNotChange, "Must-not-change list"],
  ];
  return checks.filter(([m]) => isUnresolved(m)).map(([, label]) => ({ label }));
}

export function ReviewStep({
  form,
  aiConfigured,
}: {
  form: MacroFormData;
  aiConfigured: boolean;
}) {
  const spec = buildSpecification(form);
  const summary = summarizeSpecification(spec);
  const unresolved = collectUnresolved(form);

  return (
    <div className="card">
      <h2 className="card-title">3. Review and generate</h2>
      <p className="card-subtitle">
        This is the exact structured specification that will be used to generate the macro. Nothing here is sent
        anywhere until you click Generate.
      </p>

      <div className="summary-sentence" aria-label="Plain-language summary">
        {summary}
      </div>

      {unresolved.length > 0 && (
        <div className="callout callout--amber" role="status" style={{ marginTop: 16 }}>
          <span className="callout-title">Fields left blank</span>
          These optional fields are neither filled in nor marked &quot;Not applicable&quot;. Go back and resolve them
          (fill them in, or check Not applicable) so the generated macro doesn&apos;t have to silently guess:
          <ul className="list-plain">
            {unresolved.map((u) => (
              <li key={u.label}>{u.label}</li>
            ))}
          </ul>
        </div>
      )}

      <h3 className="section-heading">Structured specification</h3>
      <dl className="kv-grid">
        <div>
          <dt>Project title</dt>
          <dd>{spec.projectTitle}</dd>
        </div>
        <div>
          <dt>Macro name</dt>
          <dd>{spec.macroName}</dd>
        </div>
        <div>
          <dt>Platform</dt>
          <dd>{spec.platform}</dd>
        </div>
        <div>
          <dt>Trigger</dt>
          <dd>{spec.execution.trigger}</dd>
        </div>
        <div>
          <dt>Source</dt>
          <dd>
            {spec.source.workbook} — {spec.source.worksheet} ({spec.source.rangeOrTable})
          </dd>
        </div>
        <div>
          <dt>Destination</dt>
          <dd>
            {spec.destination.sameWorkbook ? "Same workbook" : spec.destination.workbook} —{" "}
            {spec.destination.worksheet} ({spec.destination.rangeOrTable})
          </dd>
        </div>
        <div>
          <dt>Append / overwrite</dt>
          <dd>{spec.rules.appendOrOverwrite}</dd>
        </div>
        <div>
          <dt>Match field</dt>
          <dd>{spec.rules.matchField}</dd>
        </div>
        <div>
          <dt>Filters</dt>
          <dd>{spec.rules.filters}</dd>
        </div>
        <div>
          <dt>Transformation rules</dt>
          <dd>{spec.rules.transformationRules}</dd>
        </div>
        <div>
          <dt>Sort order</dt>
          <dd>{spec.rules.sortOrder}</dd>
        </div>
        <div>
          <dt>Duplicate handling</dt>
          <dd>{spec.rules.duplicateHandling}</dd>
        </div>
        <div>
          <dt>Blank/error handling</dt>
          <dd>{spec.rules.blankOrErrorHandling}</dd>
        </div>
        <div>
          <dt>Constraints</dt>
          <dd>{spec.execution.constraints}</dd>
        </div>
        <div>
          <dt>Must not change</dt>
          <dd>{spec.execution.mustNotChange}</dd>
        </div>
      </dl>

      {!aiConfigured && (
        <div className="callout callout--amber" role="status" style={{ marginTop: 16 }}>
          <span className="callout-title">AI generation not configured</span>
          This server has no <code>OPENAI_API_KEY</code> set, so no VBA code can be generated right now. Nothing has
          been generated or faked. You can still copy or download the structured specification above and generate
          the macro once a key is configured — see the README.
        </div>
      )}

      <div className="callout" role="note" style={{ marginTop: 16 }}>
        <span className="callout-title">Review before running</span>
        This app never opens, edits, or executes any Excel workbook, and it never verifies that these file paths,
        sheets, or tables actually exist. Always read the generated VBA yourself and test it on a copy of your
        workbook before using it on real data.
      </div>
    </div>
  );
}

export { collectUnresolved };
