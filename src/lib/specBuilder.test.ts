import { describe, expect, it } from "vitest";
import { emptyFormData, emptyMaybe } from "./types";
import { buildSpecification, summarizeSpecification } from "./specBuilder";

function fixedNow() {
  return "2026-01-01T00:00:00.000Z";
}

describe("buildSpecification", () => {
  it("carries every entered field through unmodified", () => {
    const form = emptyFormData();
    form.task.projectTitle = "Monthly Sales Rollup";
    form.task.macroName = "RollUpSales";
    form.task.platform = "windows";
    form.task.problemDescription = "Sales reps submit raw data that needs consolidating.";
    form.task.desiredResult = "One clean summary sheet per month.";
    form.task.successCondition = "Totals match the source sheet's sum.";

    form.mapping.sourceWorkbook = "Sample_Sales.xlsx";
    form.mapping.sourceWorksheet = "Raw Data";
    form.mapping.sourceRangeOrTable = "A1:F500";
    form.mapping.sourceHeaderRow = "1";
    form.mapping.sourceColumnHeaders = "Date, Region, Amount";
    form.mapping.sourceExampleRows = { notApplicable: false, value: "2026-01-01, East, 100" };

    form.mapping.sameWorkbook = false;
    form.mapping.destinationWorkbook = { notApplicable: false, value: "Monthly_Summary.xlsx" };
    form.mapping.destinationWorksheet = "Monthly Summary";
    form.mapping.destinationRangeOrTable = "A1";

    form.mapping.matchField = { notApplicable: false, value: "Region" };
    form.mapping.filters = { notApplicable: false, value: "Exclude cancelled orders" };
    form.mapping.transformationRules = { notApplicable: false, value: "Sum Amount by Region" };
    form.mapping.sortOrder = { notApplicable: false, value: "Region ascending" };
    form.mapping.duplicateHandling = { notApplicable: false, value: "Keep the last occurrence" };
    form.mapping.blankOrErrorHandling = { notApplicable: false, value: "Skip blank rows" };
    form.mapping.appendOrOverwrite = "overwrite";

    form.run.trigger = "button";
    form.run.constraints = { notApplicable: false, value: "Must finish in under 5 seconds" };
    form.run.mustNotChange = { notApplicable: false, value: "Do not touch the Raw Data sheet" };

    const spec = buildSpecification(form, fixedNow);

    expect(spec.projectTitle).toBe("Monthly Sales Rollup");
    expect(spec.macroName).toBe("RollUpSales");
    expect(spec.platform).toBe("windows");
    expect(spec.problem.description).toBe(form.task.problemDescription);
    expect(spec.source.workbook).toBe("Sample_Sales.xlsx");
    expect(spec.source.exampleRows).toBe("2026-01-01, East, 100");
    expect(spec.destination.sameWorkbook).toBe(false);
    expect(spec.destination.workbook).toBe("Monthly_Summary.xlsx");
    expect(spec.rules.matchField).toBe("Region");
    expect(spec.rules.appendOrOverwrite).toBe("overwrite");
    expect(spec.execution.trigger).toBe("button");
    expect(spec.execution.constraints).toBe("Must finish in under 5 seconds");
  });

  it("renders Not applicable fields as the literal string, not invented text", () => {
    const form = emptyFormData();
    form.mapping.matchField = { notApplicable: true, value: "" };
    form.mapping.sourceExampleRows = { notApplicable: true, value: "should be ignored" };
    const spec = buildSpecification(form, fixedNow);
    expect(spec.rules.matchField).toBe("Not applicable");
    expect(spec.source.exampleRows).toBe("Not applicable");
  });

  it("treats same-workbook destination as Not applicable regardless of stray input", () => {
    const form = emptyFormData();
    form.mapping.sameWorkbook = true;
    form.mapping.destinationWorkbook = { notApplicable: false, value: "leftover.xlsx" };
    const spec = buildSpecification(form, fixedNow);
    expect(spec.destination.workbook).toBe("Not applicable");
  });

  it("treats an empty non-N/A maybe field as Not applicable rather than an empty string", () => {
    const m = emptyMaybe();
    const form = emptyFormData();
    form.mapping.filters = m;
    const spec = buildSpecification(form, fixedNow);
    expect(spec.rules.filters).toBe("Not applicable");
  });
});

describe("summarizeSpecification", () => {
  it("produces a plain sentence including source, match, and destination", () => {
    const form = emptyFormData();
    form.mapping.sourceWorkbook = "Sample_Sales.xlsx";
    form.mapping.sourceWorksheet = "Raw Data";
    form.mapping.destinationWorksheet = "Monthly Summary";
    form.mapping.matchField = { notApplicable: false, value: "Region" };
    form.mapping.appendOrOverwrite = "append";
    const spec = buildSpecification(form, fixedNow);
    const summary = summarizeSpecification(spec);
    expect(summary).toContain("Raw Data");
    expect(summary).toContain("Region");
    expect(summary).toContain("Monthly Summary");
    expect(summary).toContain("append");
  });
});
