import { describe, expect, it } from "vitest";
import { emptyFormData } from "./types";
import { validateMacroForm } from "./formValidation";

function validForm() {
  const form = emptyFormData();
  form.task.projectTitle = "Monthly Sales Rollup";
  form.task.macroName = "RollUpSales";
  form.task.problemDescription = "Consolidate raw sales rows into a summary.";
  form.task.desiredResult = "One summary sheet per month.";
  form.task.successCondition = "Totals reconcile with source.";
  form.mapping.sourceWorkbook = "Sample_Sales.xlsx";
  form.mapping.sourceWorksheet = "Raw Data";
  form.mapping.sourceRangeOrTable = "A1:F500";
  form.mapping.sourceHeaderRow = "1";
  form.mapping.sourceColumnHeaders = "Date, Region, Amount";
  form.mapping.destinationWorksheet = "Monthly Summary";
  form.mapping.destinationRangeOrTable = "A1";
  return form;
}

describe("validateMacroForm", () => {
  it("passes for a fully filled, same-workbook task", () => {
    const result = validateMacroForm(validForm());
    expect(result.valid).toBe(true);
    expect(result.errors).toEqual([]);
  });

  it("fails when required fields are blank", () => {
    const form = emptyFormData();
    const result = validateMacroForm(form);
    expect(result.valid).toBe(false);
    expect(result.errors.length).toBeGreaterThan(0);
  });

  it("fails on an invalid macro name", () => {
    const form = validForm();
    form.task.macroName = "123 Bad Name!";
    const result = validateMacroForm(form);
    expect(result.valid).toBe(false);
    expect(result.errors.some((e) => e.field === "task.macroName")).toBe(true);
  });

  it("requires a destination workbook when source and destination differ", () => {
    const form = validForm();
    form.mapping.sameWorkbook = false;
    form.mapping.destinationWorkbook = { notApplicable: false, value: "" };
    const result = validateMacroForm(form);
    expect(result.valid).toBe(false);
    expect(result.errors.some((e) => e.field === "mapping.destinationWorkbook")).toBe(true);
  });

  it("rejects marking destination workbook Not applicable for a cross-workbook task", () => {
    const form = validForm();
    form.mapping.sameWorkbook = false;
    form.mapping.destinationWorkbook = { notApplicable: true, value: "" };
    const result = validateMacroForm(form);
    expect(result.valid).toBe(false);
  });

  it("accepts a cross-workbook task once destination workbook is filled in", () => {
    const form = validForm();
    form.mapping.sameWorkbook = false;
    form.mapping.destinationWorkbook = { notApplicable: false, value: "Monthly_Summary.xlsx" };
    const result = validateMacroForm(form);
    expect(result.valid).toBe(true);
  });

  it("allows optional fields to be marked Not applicable without forcing content", () => {
    const form = validForm();
    form.mapping.matchField = { notApplicable: true, value: "" };
    form.mapping.filters = { notApplicable: true, value: "" };
    form.run.constraints = { notApplicable: true, value: "" };
    const result = validateMacroForm(form);
    expect(result.valid).toBe(true);
  });
});
