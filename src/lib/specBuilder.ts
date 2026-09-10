import { DataMapping, MacroFormData, MacroSpecification, Maybe, PreviewConfig } from "./types";

const NOT_APPLICABLE = "Not applicable" as const;

function fromMaybe(m: Maybe): string | "Not applicable" {
  if (m.notApplicable) return NOT_APPLICABLE;
  return m.value.trim().length > 0 ? m.value.trim() : NOT_APPLICABLE;
}

function fromText(value: string): string | "Not applicable" {
  return value.trim().length > 0 ? value.trim() : NOT_APPLICABLE;
}

/**
 * Projects the preview picker into the specification.
 *
 * Deliberate omission: `sampleHeaders` / `sampleRows` never leave the
 * browser. They exist only to draw the local before/after example, they are
 * frequently pasted from a user's real workbook, and the generator does not
 * need them. Only the operation shape is included.
 */
function previewSection(preview: PreviewConfig | undefined): MacroSpecification["preview"] {
  // `preview` can legitimately be absent: older saved drafts and hand-posted
  // request bodies predate this section.
  if (!preview || preview.kind === "not-configured") {
    return {
      kind: "not-configured",
      keyColumn: NOT_APPLICABLE,
      valueColumn: NOT_APPLICABLE,
      aggregate: NOT_APPLICABLE,
      filterOperator: NOT_APPLICABLE,
      filterValue: NOT_APPLICABLE,
    };
  }
  const usesKey = preview.kind !== "copy";
  const isAggregate = preview.kind === "aggregate";
  const isFilter = preview.kind === "filter";
  return {
    kind: preview.kind,
    keyColumn: usesKey ? fromText(preview.keyColumn) : NOT_APPLICABLE,
    valueColumn: isAggregate && preview.aggregate !== "count" ? fromText(preview.valueColumn) : NOT_APPLICABLE,
    aggregate: isAggregate ? preview.aggregate : NOT_APPLICABLE,
    filterOperator: isFilter ? preview.filterOperator : NOT_APPLICABLE,
    filterValue:
      isFilter && preview.filterOperator !== "is-blank" && preview.filterOperator !== "is-not-blank"
        ? fromText(preview.filterValue)
        : NOT_APPLICABLE,
  };
}

/**
 * Projects the optional second lookup source into the specification, the
 * same way `destination.workbook` handles a same-workbook destination:
 * every field is "Not applicable" when the second source is disabled.
 */
function secondSourceSection(mapping: DataMapping): MacroSpecification["secondSource"] {
  if (!mapping.secondSourceEnabled) {
    return {
      enabled: false,
      sameWorkbookAsSource: mapping.secondSourceSameWorkbookAsSource,
      workbook: NOT_APPLICABLE,
      worksheet: NOT_APPLICABLE,
      rangeOrTable: NOT_APPLICABLE,
      headerRow: NOT_APPLICABLE,
      columnHeaders: NOT_APPLICABLE,
    };
  }
  return {
    enabled: true,
    sameWorkbookAsSource: mapping.secondSourceSameWorkbookAsSource,
    workbook: mapping.secondSourceSameWorkbookAsSource ? NOT_APPLICABLE : fromMaybe(mapping.secondSourceWorkbook),
    worksheet: fromText(mapping.secondSourceWorksheet),
    rangeOrTable: fromText(mapping.secondSourceRangeOrTable),
    headerRow: fromText(mapping.secondSourceHeaderRow),
    columnHeaders: fromText(mapping.secondSourceColumnHeaders),
  };
}

/**
 * Deterministically builds a structured specification object from the form.
 * No AI, no free-text paragraph -- this is what gets sent to the prompt
 * builder, and it is also downloadable/copyable as-is when no API key is
 * configured.
 */
export function buildSpecification(form: MacroFormData, now: () => string = () => new Date().toISOString()): MacroSpecification {
  return {
    specVersion: 2,
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
    secondSource: secondSourceSection(form.mapping),
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
    preview: previewSection(form.preview),
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
