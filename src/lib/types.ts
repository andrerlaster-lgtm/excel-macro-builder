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

export interface MacroFormData {
  task: TaskDescription;
  mapping: DataMapping;
  run: RunConfig;
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
  };
}

// ---------------------------------------------------------------------------
// Structured specification: the deterministic object built from the form,
// independent of any particular AI provider or prompt phrasing.
// ---------------------------------------------------------------------------

export interface MacroSpecification {
  specVersion: 1;
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
}

// ---------------------------------------------------------------------------
// AI response, parsed defensively from provider output.
// ---------------------------------------------------------------------------

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
