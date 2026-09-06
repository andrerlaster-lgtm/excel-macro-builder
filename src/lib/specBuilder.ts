import { MacroFormData, MacroSpecification, Maybe } from "./types";

const NOT_APPLICABLE = "Not applicable" as const;

function fromMaybe(m: Maybe): string | "Not applicable" {
  if (m.notApplicable) return NOT_APPLICABLE;
  return m.value.trim().length > 0 ? m.value.trim() : NOT_APPLICABLE;
}

/**
 * Deterministically builds a structured specification object from the form.
 * No AI, no free-text paragraph -- this is what gets sent to the prompt
 * builder, and it is also downloadable/copyable as-is when no API key is
 * configured.
 */
export function buildSpecification(form: MacroFormData, now: () => string = () => new Date().toISOString()): MacroSpecification {
  return {
    specVersion: 1,
    generatedAt: now(),
    projectTitle: form.task.projectTitle.trim(),
    macroName: form.task.macroName.trim(),
    platform: form.task.platform,
    problem: {
      description: form.task.problemDescription.trim(),
      desiredResult: form.task.desiredResult.trim(),
      successCondition: form.task.successCondition.trim(),
    },
    source: {
      workbook: form.mapping.sourceWorkbook.trim(),
      worksheet: form.mapping.sourceWorksheet.trim(),
      rangeOrTable: form.mapping.sourceRangeOrTable.trim(),
      headerRow: form.mapping.sourceHeaderRow.trim(),
      columnHeaders: form.mapping.sourceColumnHeaders.trim(),
      exampleRows: fromMaybe(form.mapping.sourceExampleRows),
    },
    destination: {
      sameWorkbook: form.mapping.sameWorkbook,
      workbook: form.mapping.sameWorkbook ? NOT_APPLICABLE : fromMaybe(form.mapping.destinationWorkbook),
      worksheet: form.mapping.destinationWorksheet.trim(),
      rangeOrTable: form.mapping.destinationRangeOrTable.trim(),
    },
    rules: {
      matchField: fromMaybe(form.mapping.matchField),
      filters: fromMaybe(form.mapping.filters),
      transformationRules: fromMaybe(form.mapping.transformationRules),
      sortOrder: fromMaybe(form.mapping.sortOrder),
      duplicateHandling: fromMaybe(form.mapping.duplicateHandling),
      blankOrErrorHandling: fromMaybe(form.mapping.blankOrErrorHandling),
      appendOrOverwrite: form.mapping.appendOrOverwrite,
    },
    execution: {
      trigger: form.run.trigger,
      constraints: fromMaybe(form.run.constraints),
      mustNotChange: fromMaybe(form.run.mustNotChange),
    },
  };
}

/** Plain-sentence summary shown on the review step, before generation. */
export function summarizeSpecification(spec: MacroSpecification): string {
  const source = spec.destination.sameWorkbook
    ? `"${spec.source.worksheet}" in ${spec.source.workbook}`
    : `"${spec.source.worksheet}" in ${spec.source.workbook}`;
  const destination = spec.destination.sameWorkbook
    ? `"${spec.destination.worksheet}" in the same workbook`
    : `"${spec.destination.worksheet}" in ${spec.destination.workbook}`;
  const matchPart =
    spec.rules.matchField !== NOT_APPLICABLE ? `, match by ${spec.rules.matchField}` : "";
  const rulesPart =
    spec.rules.transformationRules !== NOT_APPLICABLE
      ? `, apply ${spec.rules.transformationRules}`
      : "";
  const modePart =
    spec.rules.appendOrOverwrite === "not-applicable"
      ? ""
      : ` using ${spec.rules.appendOrOverwrite}`;

  return `Read rows from ${source}${matchPart}${rulesPart}, and write to ${destination}${modePart}.`;
}
