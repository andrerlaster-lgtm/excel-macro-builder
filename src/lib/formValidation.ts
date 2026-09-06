import { MacroFormData, Maybe } from "./types";
import { validateVbaMacroName } from "./vbaValidation";

export const MAX_SHORT_FIELD_LENGTH = 200;
export const MAX_LONG_FIELD_LENGTH = 4000;

export interface FieldError {
  field: string;
  message: string;
}

export interface FormValidationResult {
  valid: boolean;
  errors: FieldError[];
}

function req(value: string, field: string, label: string, maxLen = MAX_SHORT_FIELD_LENGTH): FieldError[] {
  const errors: FieldError[] = [];
  if (!value || value.trim().length === 0) {
    errors.push({ field, message: `${label} is required.` });
  } else if (value.length > maxLen) {
    errors.push({ field, message: `${label} must be ${maxLen} characters or fewer.` });
  }
  return errors;
}

function maybeLen(m: Maybe, field: string, label: string, maxLen = MAX_LONG_FIELD_LENGTH): FieldError[] {
  if (m.notApplicable) return [];
  if (m.value.length > maxLen) {
    return [{ field, message: `${label} must be ${maxLen} characters or fewer.` }];
  }
  return [];
}

/**
 * Validates the full macro form. Used both client-side (progressive,
 * per-step) and server-side (defense in depth before any AI call).
 */
export function validateMacroForm(form: MacroFormData): FormValidationResult {
  const errors: FieldError[] = [];

  errors.push(...req(form.task.projectTitle, "task.projectTitle", "Project title"));

  const nameCheck = validateVbaMacroName(form.task.macroName);
  if (!nameCheck.valid) {
    for (const message of nameCheck.errors) {
      errors.push({ field: "task.macroName", message });
    }
  }

  errors.push(
    ...req(form.task.problemDescription, "task.problemDescription", "Problem description", MAX_LONG_FIELD_LENGTH)
  );
  errors.push(...req(form.task.desiredResult, "task.desiredResult", "Desired result", MAX_LONG_FIELD_LENGTH));
  errors.push(
    ...req(form.task.successCondition, "task.successCondition", "Success condition", MAX_LONG_FIELD_LENGTH)
  );

  errors.push(...req(form.mapping.sourceWorkbook, "mapping.sourceWorkbook", "Source workbook name/path"));
  errors.push(...req(form.mapping.sourceWorksheet, "mapping.sourceWorksheet", "Source worksheet"));
  errors.push(...req(form.mapping.sourceRangeOrTable, "mapping.sourceRangeOrTable", "Source range or table"));
  errors.push(...req(form.mapping.sourceHeaderRow, "mapping.sourceHeaderRow", "Source header row"));
  errors.push(
    ...req(form.mapping.sourceColumnHeaders, "mapping.sourceColumnHeaders", "Relevant column headers", MAX_LONG_FIELD_LENGTH)
  );
  errors.push(...maybeLen(form.mapping.sourceExampleRows, "mapping.sourceExampleRows", "Example rows"));

  if (!form.mapping.sameWorkbook) {
    if (form.mapping.destinationWorkbook.notApplicable) {
      errors.push({
        field: "mapping.destinationWorkbook",
        message: "Destination workbook cannot be Not applicable when source and destination are different workbooks.",
      });
    } else {
      errors.push(
        ...req(form.mapping.destinationWorkbook.value, "mapping.destinationWorkbook", "Destination workbook name/path")
      );
    }
  }

  errors.push(...req(form.mapping.destinationWorksheet, "mapping.destinationWorksheet", "Destination worksheet"));
  errors.push(
    ...req(form.mapping.destinationRangeOrTable, "mapping.destinationRangeOrTable", "Destination range/table or starting cell")
  );

  errors.push(...maybeLen(form.mapping.matchField, "mapping.matchField", "Match/lookup field"));
  errors.push(...maybeLen(form.mapping.filters, "mapping.filters", "Filters"));
  errors.push(...maybeLen(form.mapping.transformationRules, "mapping.transformationRules", "Transformation/calculation rules"));
  errors.push(...maybeLen(form.mapping.sortOrder, "mapping.sortOrder", "Sort order"));
  errors.push(...maybeLen(form.mapping.duplicateHandling, "mapping.duplicateHandling", "Duplicate handling"));
  errors.push(...maybeLen(form.mapping.blankOrErrorHandling, "mapping.blankOrErrorHandling", "Blank/error handling"));

  errors.push(...maybeLen(form.run.constraints, "run.constraints", "Constraints"));
  errors.push(...maybeLen(form.run.mustNotChange, "run.mustNotChange", "Must-not-change list"));

  return { valid: errors.length === 0, errors };
}
