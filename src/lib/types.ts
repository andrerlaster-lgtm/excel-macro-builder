// Core data model for Excel Macro Builder.
//
// This app never opens, uploads, or executes a workbook. Everything here is
// just text the user typed, structured enough to build a deterministic
// specification and prompt for the AI provider adapter.

export type ExcelPlatform = "windows" | "mac" | "cross-platform";

export type TriggerMode = "manual" | "button" | "workbook-open" | "sheet-change";

export type AppendOrOverwrite = "append" | "overwrite" | "not-applicable";

/** A field that can be filled in, or explicitly marked N/A instead of guessed. */
export interface Maybe {
  /** true = user marked this "Not applicable" on purpose. */
  notApplicable: boolean;
  value: string;
}

export function emptyMaybe(): Maybe {
  return { notApplicable: false, value: "" };
}

export interface TaskDescription {
  projectTitle: string;
  macroName: string;
  platform: ExcelPlatform;
  problemDescription: string;
  desiredResult: string;
  successCondition: string;
}

export interface DataMapping {
  sourceWorkbook: string;
  sourceWorksheet: string;
  sourceRangeOrTable: string;
  sourceHeaderRow: string;
  sourceColumnHeaders: string;
  sourceExampleRows: Maybe;

  sameWorkbook: boolean;
  destinationWorkbook: Maybe;
  destinationWorksheet: string;
  destinationRangeOrTable: string;

  matchField: Maybe;
  filters: Maybe;
  transformationRules: Maybe;
  sortOrder: Maybe;
  duplicateHandling: Maybe;
  blankOrErrorHandling: Maybe;
  appendOrOverwrite: AppendOrOverwrite;
}

export interface RunConfig {
  trigger: TriggerMode;
  constraints: Maybe;
  mustNotChange: Maybe;
}

// ---------------------------------------------------------------------------
// Preview configuration.
//
// This describes a small, structured version of the operation the user says
// they want, purely so the app can compute a deterministic before/after
// example in the browser. It is NOT what generates the VBA -- the free-text
// rule fields in DataMapping remain the source of truth for generation, and
// the preview never executes or inspects the generated code.
// ---------------------------------------------------------------------------

export type PreviewOperationKind = "not-configured" | "copy" | "filter" | "aggregate" | "deduplicate" | "lookup";

export type AggregateFn = "sum" | "count" | "average" | "min" | "max";

export type FilterOperator =
  | "equals"
  | "not-equals"
  | "contains"
  | "greater-than"
  | "less-than"
  | "is-blank"
  | "is-not-blank";

export interface PreviewConfig {
  kind: PreviewOperationKind;
  /** group-by column (aggregate/deduplicate) or column under test (filter) */
  keyColumn: string;
  /** numeric column to aggregate */
  valueColumn: string;
  aggregate: AggregateFn;
  filterOperator: FilterOperator;
  filterValue: string;
  /** Editable sample grid used by the preview. Seeded from mapping.sourceExampleRows. */
  sampleHeaders: string[];
  sampleRows: string[][];
  /**
   * Editable DESTINATION sample grid, used only by "lookup" (kind === "lookup").
   * Present but unused for the other four kinds -- this type stays a flat
   * interface rather than a discriminated union, matching the rest of this
   * file's style. Lookup is the only operation that reads existing
   * destination rows before writing, so it needs its own sample table
   * distinct from `sampleHeaders`/`sampleRows`, which are always the SOURCE.
   */
  destSampleHeaders: string[];
  destSampleRows: string[][];
}

export function emptyPreviewConfig(): PreviewConfig {
  return {
    kind: "not-configured",
    keyColumn: "",
    valueColumn: "",
    aggregate: "sum",
    filterOperator: "equals",
    filterValue: "",
    sampleHeaders: [],
    sampleRows: [],
    destSampleHeaders: [],
    destSampleRows: [],
  };
}

export interface MacroFormData {
  task: TaskDescription;
  mapping: DataMapping;
  run: RunConfig;
  preview: PreviewConfig;
}

export function emptyFormData(): MacroFormData {
  return {
    task: {
      projectTitle: "",
      macroName: "",
      platform: "cross-platform",
      problemDescription: "",
      desiredResult: "",
      successCondition: "",
    },
    mapping: {
      sourceWorkbook: "",
      sourceWorksheet: "",
      sourceRangeOrTable: "",
      sourceHeaderRow: "",
      sourceColumnHeaders: "",
      sourceExampleRows: emptyMaybe(),
      sameWorkbook: true,
      destinationWorkbook: emptyMaybe(),
      destinationWorksheet: "",
      destinationRangeOrTable: "",
      matchField: emptyMaybe(),
      filters: emptyMaybe(),
      transformationRules: emptyMaybe(),
      sortOrder: emptyMaybe(),
      duplicateHandling: emptyMaybe(),
      blankOrErrorHandling: emptyMaybe(),
      appendOrOverwrite: "overwrite",
    },
    run: {
      trigger: "manual",
      constraints: emptyMaybe(),
      mustNotChange: emptyMaybe(),
    },
    preview: emptyPreviewConfig(),
  };
}

// ---------------------------------------------------------------------------
// Structured specification: the deterministic object built from the form,
// independent of any particular AI provider or prompt phrasing.
// ---------------------------------------------------------------------------

export interface MacroSpecification {
  /** 1 = original shape; 2 added the `preview` section. */
  specVersion: 2;
  generatedAt: string;
  projectTitle: string;
  macroName: string;
  platform: ExcelPlatform;
  problem: {
    description: string;
    desiredResult: string;
    successCondition: string;
  };
  source: {
    workbook: string;
    worksheet: string;
    rangeOrTable: string;
    headerRow: string;
    columnHeaders: string;
    exampleRows: string | "Not applicable";
  };
  destination: {
    sameWorkbook: boolean;
    workbook: string | "Not applicable";
    worksheet: string;
    rangeOrTable: string;
  };
  rules: {
    matchField: string | "Not applicable";
    filters: string | "Not applicable";
    transformationRules: string | "Not applicable";
    sortOrder: string | "Not applicable";
    duplicateHandling: string | "Not applicable";
    blankOrErrorHandling: string | "Not applicable";
    appendOrOverwrite: AppendOrOverwrite;
  };
  execution: {
    trigger: TriggerMode;
    constraints: string | "Not applicable";
    mustNotChange: string | "Not applicable";
  };
  /**
   * The shape of the operation the user picked for the in-browser preview.
   * Deliberately carries NO sample rows: those are illustrative, may be
   * pasted from real data, and are never worth sending to an AI provider.
   * Only the operation shape travels.
   */
  preview: {
    kind: PreviewOperationKind;
    keyColumn: string | "Not applicable";
    valueColumn: string | "Not applicable";
    aggregate: AggregateFn | "Not applicable";
    filterOperator: FilterOperator | "Not applicable";
    filterValue: string | "Not applicable";
  };
}

// ---------------------------------------------------------------------------
// Generated macro result.
//
// Two generators can produce this shape: the AI provider (parsed defensively
// from provider output) and the deterministic in-browser template builder in
// `src/lib/vbaTemplateGenerator.ts`. Which one produced a result is recorded
// on it so a reopened project can never misattribute template code as AI
// output, or vice versa.
// ---------------------------------------------------------------------------

export type GeneratorKind = "ai" | "template";

export interface MacroGenerationResult {
  vbaCode: string;
  summary: string;
  assumptions: string[];
  openQuestions: string[];
  inputsOutputs: string;
  installInstructions: string;
  testPlan: string[];
  safetyCautions: string[];
  platformLimitations: string[];
  /**
   * What the emitted code does, in order. Only the deterministic template can
   * state this honestly, so it is absent on AI results.
   */
  steps?: string[];
  /**
   * Which generator produced this result. Optional because records saved
   * before schema v3 predate the field; the storage migration backfills
   * those to "ai", which is the only generator that existed then.
   */
  generator?: GeneratorKind;
  /**
   * Free-text rules the deterministic template did NOT encode. Structurally
   * matches `UnimplementedRule` in `src/lib/vbaTemplateGenerator.ts` (inlined
   * here to keep this module free of a dependency on the generator).
   * Template results only.
   */
  unimplementedRules?: { field: string; text: string }[];
}

// ---------------------------------------------------------------------------
// Saved projects (versioned localStorage records).
// ---------------------------------------------------------------------------

export interface SavedProject {
  id: string;
  schemaVersion: number;
  title: string;
  createdAt: string;
  updatedAt: string;
  archived: boolean;
  form: MacroFormData;
  lastSpecification: MacroSpecification | null;
  lastResult: MacroGenerationResult | null;
}
