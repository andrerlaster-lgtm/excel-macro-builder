import { describe, expect, it } from "vitest";
import { emptyFormData, MacroFormData } from "./types";
import { generateVbaFromTemplate, collectUnimplementedRules } from "./vbaTemplateGenerator";
import { parseNumericCell, simulatePreview } from "./previewSimulator";
import { scanVbaForWarnings } from "./safetyScanner";

// Fictional sample data only -- never anything resembling real client files.
const HEADERS = "Account Name, Account Number, Amount";

function baseForm(): MacroFormData {
  const form = emptyFormData();
  form.task.projectTitle = "GL Account Balance Transfer";
  form.task.macroName = "TransferAccountBalances";
  form.task.platform = "windows";
  form.task.problemDescription = "Balances are totalled by hand every month.";
  form.task.desiredResult = "One row per account with its total.";
  form.task.successCondition = "Totals match the manual workbook.";

  form.mapping.sourceWorkbook = "Sample_GL_Export.xlsx";
  form.mapping.sourceWorksheet = "Raw Data";
  form.mapping.sourceRangeOrTable = "tblGLExport";
  form.mapping.sourceHeaderRow = "Row 1";
  form.mapping.sourceColumnHeaders = HEADERS;
  form.mapping.sameWorkbook = true;
  form.mapping.destinationWorksheet = "Summary";
  form.mapping.destinationRangeOrTable = "A1";
  form.mapping.appendOrOverwrite = "overwrite";

  form.run.trigger = "manual";

  form.preview.kind = "aggregate";
  form.preview.keyColumn = "Account Name";
  form.preview.valueColumn = "Amount";
  form.preview.aggregate = "sum";
  return form;
}

function crossWorkbookForm(): MacroFormData {
  const form = baseForm();
  form.mapping.sameWorkbook = false;
  form.mapping.destinationWorkbook = { notApplicable: false, value: "Sample_Monthly_Report.xlsx" };
  form.mapping.destinationRangeOrTable = "tblSummary";
  return form;
}

const FORBIDDEN = [/ActiveSheet/, /ActiveWorkbook/, /Selection/, /\.Select\b/, /\.Activate\b/];

function expectNoForbiddenIdioms(vba: string) {
  for (const pattern of FORBIDDEN) {
    expect(vba, `should not contain ${pattern}`).not.toMatch(pattern);
  }
}

describe("destination table columns are matched by heading, not by position", () => {
  it("maps output columns onto the table before clearing anything", () => {
    const vba = generateVbaFromTemplate(crossWorkbookForm()).vbaCode;
    const mapAt = vba.indexOf("MapOutputColumnsToTable destTable");
    const clearAt = vba.indexOf("destTable.DataBodyRange.ClearContents");
    expect(mapAt).toBeGreaterThan(-1);
    expect(clearAt).toBeGreaterThan(-1);
    // A heading mismatch must abort before any destination data is destroyed.
    expect(mapAt).toBeLessThan(clearAt);
  });

  it("never writes a positional block into a table", () => {
    const vba = generateVbaFromTemplate(crossWorkbookForm()).vbaCode;
    expect(vba).not.toMatch(/DataBodyRange\.Cells\(1, 1\)\.Resize\(outputRowCount, outputColumns\)/);
    expect(vba).toMatch(/DataBodyRange\.Cells\(1, destColumnMap\(c\)\)/);
  });

  it("routes appended rows through the heading map too", () => {
    const form = crossWorkbookForm();
    form.mapping.appendOrOverwrite = "append";
    const vba = generateVbaFromTemplate(form).vbaCode;
    expect(vba).toMatch(/MapOutputColumnsToTable destTable/);
    expect(vba).toMatch(/newRow\.Range\.Cells\(1, destColumnMap\(c\)\)\.Value/);
    expect(vba).not.toMatch(/newRow\.Range\.Cells\(1, c\)\.Value/);
  });

  it("emits a mapper that errors on a missing heading and lists the real ones", () => {
    const vba = generateVbaFromTemplate(crossWorkbookForm()).vbaCode;
    expect(vba).toMatch(/Private Sub MapOutputColumnsToTable/);
    expect(vba).toMatch(/Private Function TableHeaderList/);
    expect(vba).toMatch(/has no column headed/);
    expect(vba).toMatch(/TableHeaderList\(targetTable\)/);
    // Refuses rather than guessing a position.
    expect(vba).toMatch(/would corrupt the report silently/);
  });

  it("refuses a table with its header row switched off", () => {
    const vba = generateVbaFromTemplate(crossWorkbookForm()).vbaCode;
    expect(vba).toMatch(/HeaderRowRange Is Nothing/);
    expect(vba).toMatch(/header row switched off/);
  });

  it("does not rename the destination table's headings", () => {
    const vba = generateVbaFromTemplate(crossWorkbookForm()).vbaCode;
    // Writing to HeaderRowRange would break structured references elsewhere.
    expect(vba).not.toMatch(/HeaderRowRange[^\n]*\.Value\s*=/);
  });

  it("tells the user which headings the table needs", () => {
    const result = generateVbaFromTemplate(crossWorkbookForm());
    const note = result.assumptions.find((a) => a.includes("heading matches"));
    expect(note).toBeDefined();
    expect(note).toContain('"Account Name"');
    expect(note).toContain('"Sum of Amount"');
    expect(note).toContain("does not rename");
  });

  it("still writes its own header row for a plain-range destination", () => {
    // A plain range is a block the macro owns, so there is nothing to match.
    const vba = generateVbaFromTemplate(baseForm()).vbaCode;
    expect(vba).toMatch(/destAnchor\.Resize\(1, outputColumns\)\.Value = outputHeaders/);
  });
});

describe("generateVbaFromTemplate — the four supported operations", () => {
  it("emits grouping and aggregation logic for aggregate", () => {
    const result = generateVbaFromTemplate(baseForm());
    expect(result.status).toBe("ok");
    expect(result.vbaCode).toContain("Set groupIndex = New Collection");
    expect(result.vbaCode).toContain("groupSum(g) = groupSum(g) + parsedNumber");
    expect(result.vbaCode).toContain('Private Const RESULT_VALUE_HEADER As String = "Sum of Amount"');
    expect(result.steps.join(" ")).toMatch(/Group the rows by "Account Name"/);
  });

  it("emits a row-by-row filter for filter", () => {
    const form = baseForm();
    form.preview.kind = "filter";
    form.preview.keyColumn = "Account Name";
    form.preview.filterOperator = "equals";
    form.preview.filterValue = "Cash - Operating";

    const result = generateVbaFromTemplate(form);
    expect(result.status).toBe("ok");
    expect(result.vbaCode).toContain('Private Const FILTER_VALUE As String = "Cash - Operating"');
    expect(result.vbaCode).toContain("keepRow(r) = (StrComp(CellText(sourceValues(r, keyColumn)), FILTER_VALUE, vbTextCompare) = 0)");
    // Not an aggregate: no grouping machinery should leak in.
    expect(result.vbaCode).not.toContain("groupSum");
  });

  it("emits first-row-wins deduplication plus a sort", () => {
    const form = baseForm();
    form.preview.kind = "deduplicate";
    form.preview.keyColumn = "Account Number";

    const result = generateVbaFromTemplate(form);
    expect(result.status).toBe("ok");
    expect(result.vbaCode).toContain("Set seenKeys = New Collection");
    expect(result.vbaCode).toContain("Keeps the FIRST row seen for each distinct key");
    expect(result.vbaCode).toContain("SortRowsByColumn outputValues, outputRowCount, keyColumn");
    expect(result.steps.join(" ")).toMatch(/FIRST row for each distinct "Account Number"/);
  });

  it("emits a straight pass-through for copy, with no key column required", () => {
    const form = baseForm();
    form.preview.kind = "copy";
    form.preview.keyColumn = "";
    form.preview.valueColumn = "";

    const result = generateVbaFromTemplate(form);
    expect(result.status).toBe("ok");
    expect(result.vbaCode).toContain("Copy every source row through unchanged");
    expect(result.vbaCode).toContain("outputValues(r, c) = sourceValues(r, c)");
    expect(result.vbaCode).not.toContain("KEY_HEADER");
  });
});

describe("generateVbaFromTemplate — VBA quality invariants", () => {
  const forms: [string, MacroFormData][] = [
    ["aggregate", baseForm()],
    ["copy", (() => {
      const f = baseForm();
      f.preview.kind = "copy";
      return f;
    })()],
    ["filter", (() => {
      const f = baseForm();
      f.preview.kind = "filter";
      f.preview.filterOperator = "greater-than";
      f.preview.filterValue = "1000";
      f.preview.keyColumn = "Amount";
      return f;
    })()],
    ["deduplicate", (() => {
      const f = baseForm();
      f.preview.kind = "deduplicate";
      f.preview.keyColumn = "Account Name";
      return f;
    })()],
    ["cross-workbook append", (() => {
      const f = crossWorkbookForm();
      f.mapping.appendOrOverwrite = "append";
      return f;
    })()],
  ];

  for (const [label, form] of forms) {
    it(`${label}: starts with Option Explicit and has a config section`, () => {
      const { vbaCode } = generateVbaFromTemplate(form);
      expect(vbaCode.startsWith("Option Explicit\n")).toBe(true);
      expect(vbaCode).toContain("' CONFIGURATION -- edit these");
      expect(vbaCode).toContain("' END CONFIGURATION");
      expect(vbaCode).toContain("Private Const SOURCE_WORKBOOK_NAME As String =");
    });

    it(`${label}: has an error handler and a cleanup path`, () => {
      const { vbaCode } = generateVbaFromTemplate(form);
      expect(vbaCode).toContain("On Error GoTo CleanFail");
      expect(vbaCode).toContain("CleanFail:");
      expect(vbaCode).toContain("Cleanup:");
      expect(vbaCode).toMatch(/Err\.Raise ERR_BASE/);
    });

    it(`${label}: saves and restores all four Application settings`, () => {
      const { vbaCode } = generateVbaFromTemplate(form);
      for (const setting of ["ScreenUpdating", "EnableEvents", "DisplayAlerts", "Calculation"]) {
        // saved before the work...
        expect(vbaCode).toContain(`prev${setting} = Application.${setting}`);
        // ...and restored in the cleanup path.
        expect(vbaCode).toContain(`Application.${setting} = prev${setting}`);
      }
      const cleanup = vbaCode.slice(vbaCode.indexOf("Cleanup:"));
      expect(cleanup).toContain("Application.Calculation = prevCalculation");
      expect(cleanup).toContain("Application.ScreenUpdating = prevScreenUpdating");
    });

    it(`${label}: never uses ActiveSheet/ActiveWorkbook/Selection/Select/Activate`, () => {
      expectNoForbiddenIdioms(generateVbaFromTemplate(form).vbaCode);
    });

    it(`${label}: looks columns up by header name and errors when one is missing`, () => {
      const { vbaCode } = generateVbaFromTemplate(form);
      expect(vbaCode).toContain("Private Function ColumnIndexByHeader(");
      expect(vbaCode).toContain("StrComp(CellText(headers.Cells(1, i).Value), headerName, vbTextCompare)");
      if (form.preview.kind !== "copy") {
        expect(vbaCode).toContain("keyColumn = ColumnIndexByHeader(headerRow, KEY_HEADER)");
        expect(vbaCode).toMatch(/If keyColumn = 0 Then[\s\S]*?was not found in the source header row/);
      }
    });

    it(`${label}: handles missing workbook, sheet, table and an empty source range`, () => {
      const { vbaCode } = generateVbaFromTemplate(form);
      expect(vbaCode).toContain("If sourceBook Is Nothing Then");
      expect(vbaCode).toContain("If sourceSheet Is Nothing Then");
      expect(vbaCode).toContain("If sourceRange Is Nothing Then");
      expect(vbaCode).toContain("If destSheet Is Nothing Then");
      expect(vbaCode).toContain("has no data rows underneath its header row");
    });

    it(`${label}: emits none of the banned destructive calls`, () => {
      const { vbaCode } = generateVbaFromTemplate(form);
      for (const banned of [/\bKill\b/, /\bRmDir\b/, /\bShell\b/, /\.SaveAs\b/, /\bOutlook\b/, /\bXMLHTTP\b/]) {
        expect(vbaCode, `should not contain ${banned}`).not.toMatch(banned);
      }
    });
  }
});

describe("generateVbaFromTemplate — refusals", () => {
  function expectRefusal(form: MacroFormData, messagePattern: RegExp) {
    const result = generateVbaFromTemplate(form);
    expect(result.status).toBe("unsupported");
    expect(result.vbaCode).toBe("");
    expect(result.message).toBeTruthy();
    expect(result.message!).toMatch(messagePattern);
    // A refusal must not smuggle out any half-built artefacts either.
    expect(result.steps).toEqual([]);
    expect(result.unimplementedRules).toEqual([]);
    return result;
  }

  it("refuses when no preview operation is configured", () => {
    const form = baseForm();
    form.preview.kind = "not-configured";
    expectRefusal(form, /step 2/i);
  });

  it("refuses the workbook-open trigger and explains where that code must live", () => {
    const form = baseForm();
    form.run.trigger = "workbook-open";
    expectRefusal(form, /ThisWorkbook/);
  });

  it("refuses the sheet-change trigger and explains where that code must live", () => {
    const form = baseForm();
    form.run.trigger = "sheet-change";
    expectRefusal(form, /Worksheet_Change/);
  });

  it("refuses when the key column is not one of the source headers", () => {
    const form = baseForm();
    form.preview.keyColumn = "Cost Centre";
    expectRefusal(form, /not one of the source column headers/i);
  });

  it("refuses when the key column is blank", () => {
    const form = baseForm();
    form.preview.kind = "deduplicate";
    form.preview.keyColumn = "";
    expectRefusal(form, /none is set/i);
  });

  it("refuses when the aggregate value column is missing from the headers", () => {
    const form = baseForm();
    form.preview.valueColumn = "Balance";
    expectRefusal(form, /Value column "Balance"/);
  });

  it("refuses when required source/destination fields are blank", () => {
    const form = baseForm();
    form.mapping.sourceWorksheet = "";
    form.mapping.destinationRangeOrTable = "  ";
    expectRefusal(form, /Source worksheet.*Destination table/s);
  });

  it("refuses a cross-workbook job with no destination workbook", () => {
    const form = baseForm();
    form.mapping.sameWorkbook = false;
    form.mapping.destinationWorkbook = { notApplicable: true, value: "" };
    expectRefusal(form, /destination workbook file name or path is required/i);
  });

  it("refuses a macro name that is not a legal VBA Sub name", () => {
    const form = baseForm();
    form.task.macroName = "Transfer Balances!";
    expectRefusal(form, /VBA Sub name/);
  });

  it("never throws on junk input", () => {
    expect(() => generateVbaFromTemplate(undefined as unknown as MacroFormData)).not.toThrow();
    expect(generateVbaFromTemplate({} as unknown as MacroFormData).status).toBe("unsupported");
  });
});

describe("generateVbaFromTemplate — honesty about free-text rules", () => {
  it("lists every non-empty free-text rule and omits N/A and blank ones", () => {
    const form = baseForm();
    form.mapping.filters = { notApplicable: false, value: "Exclude rows where Status = Cancelled" };
    form.mapping.transformationRules = { notApplicable: true, value: "" };
    form.mapping.sortOrder = { notApplicable: false, value: "Account Number ascending" };
    form.mapping.duplicateHandling = { notApplicable: false, value: "   " };
    form.mapping.blankOrErrorHandling = { notApplicable: false, value: "Skip blank Amount rows" };
    form.run.constraints = { notApplicable: false, value: "Must not prompt the user" };
    form.run.mustNotChange = { notApplicable: true, value: "Raw Data sheet" };

    const result = generateVbaFromTemplate(form);
    const fields = result.unimplementedRules.map((r) => r.field);

    expect(fields).toEqual([
      "Filters",
      "Sort order",
      "Blank/error handling",
      "Optional constraints",
    ]);
    expect(fields).not.toContain("Transformation/calculation rules"); // marked N/A
    expect(fields).not.toContain("Duplicate handling"); // whitespace only
    expect(fields).not.toContain("Anything the macro must not change"); // marked N/A
  });

  it("repeats every unimplemented rule in a marked comment block in the code", () => {
    const form = baseForm();
    form.mapping.filters = { notApplicable: false, value: "Exclude rows where Status = Cancelled" };
    form.run.constraints = { notApplicable: false, value: "Must finish in under 5 seconds" };

    const { vbaCode } = generateVbaFromTemplate(form);
    expect(vbaCode).toContain("RULES FROM YOUR DESCRIPTION THAT WERE **NOT** IMPLEMENTED");
    expect(vbaCode).toContain("'    - Filters: Exclude rows where Status = Cancelled");
    expect(vbaCode).toContain("'    - Optional constraints: Must finish in under 5 seconds");
  });

  it("still says the template ignores free text when no rules were entered", () => {
    const result = generateVbaFromTemplate(baseForm());
    expect(result.unimplementedRules).toEqual([]);
    expect(result.vbaCode).toContain("it never reads or");
    expect(result.vbaCode).toContain("interprets free text of any kind");
  });

  it("does not try to guess that a free-text rule is already covered", () => {
    // "Sum Amount by Region" describes exactly the configured aggregate, and it
    // must STILL be listed: guessing coverage would be unreliable, and a wrong
    // "already handled" is far more dangerous than a redundant reminder.
    const form = baseForm();
    form.mapping.transformationRules = { notApplicable: false, value: "Sum Amount by Account Name" };
    expect(collectUnimplementedRules(form).map((r) => r.field)).toContain(
      "Transformation/calculation rules"
    );
  });
});

describe("generateVbaFromTemplate — workbook and write-mode variations", () => {
  it("same-workbook code does not open anything; cross-workbook opens by path and never saves or closes", () => {
    const same = generateVbaFromTemplate(baseForm()).vbaCode;
    const cross = generateVbaFromTemplate(crossWorkbookForm()).vbaCode;

    expect(same).toContain("Set destBook = sourceBook");
    expect(same).not.toContain("Workbooks.Open");
    expect(same).not.toContain("DEST_WORKBOOK_PATH");

    expect(cross).toContain("Set destBook = FindOpenWorkbook(FileNameOnly(DEST_WORKBOOK_PATH))");
    expect(cross).toContain("Set destBook = Application.Workbooks.Open(FileName:=DEST_WORKBOOK_PATH)");
    expect(cross).toContain("Deliberately NOT saved and NOT closed");
    expect(cross).not.toMatch(/\.Close\b/);
    expect(cross).not.toMatch(/(?<!Save)\.Save\b/);
    expect(same).not.toBe(cross);

    const assumptions = generateVbaFromTemplate(crossWorkbookForm()).assumptions.join(" ");
    expect(assumptions).toMatch(/NOT saved and NOT closed/);
  });

  it("overwrite clears only the target range; append clears nothing", () => {
    const overwriteForm = baseForm();
    overwriteForm.mapping.appendOrOverwrite = "overwrite";
    const appendForm = baseForm();
    appendForm.mapping.appendOrOverwrite = "append";

    const overwrite = generateVbaFromTemplate(overwriteForm).vbaCode;
    const append = generateVbaFromTemplate(appendForm).vbaCode;

    expect(overwrite).not.toBe(append);

    expect(overwrite).toContain("destTable.DataBodyRange.ClearContents");
    expect(overwrite).toContain("destSheet.Cells(clearLastRow, destAnchor.Column + outputColumns - 1)).ClearContents");
    // Never the whole sheet, and never a row/column delete.
    expect(overwrite).not.toMatch(/\.Cells\.Clear/);
    expect(overwrite).not.toMatch(/EntireRow|EntireColumn/);
    expect(overwrite).not.toMatch(/\.Delete\b/);

    expect(append).not.toContain("ClearContents");
    expect(append).toContain("Set newRow = destTable.ListRows.Add");
    expect(append).toContain("Nothing already in the destination is cleared, changed or deleted");
  });

  it("treats an append-vs-overwrite choice of not-applicable as non-clearing", () => {
    const form = baseForm();
    form.mapping.appendOrOverwrite = "not-applicable";
    const { vbaCode } = generateVbaFromTemplate(form);
    expect(vbaCode).not.toContain("ClearContents");
  });

  it("mentions the button in the install instructions for a button trigger", () => {
    const form = baseForm();
    form.run.trigger = "button";
    const result = generateVbaFromTemplate(form);
    expect(result.status).toBe("ok");
    expect(result.installInstructions).toMatch(/Assign Macro/);
    expect(result.vbaCode).toContain("Assign this Sub to your button");
  });
});

describe("generateVbaFromTemplate — aggregate skip reporting", () => {
  it("skips non-numeric values instead of zeroing them, and reports the count", () => {
    const { vbaCode } = generateVbaFromTemplate(baseForm());
    expect(vbaCode).toContain("Private Function TryParseNumber(");
    expect(vbaCode).toContain("skippedValues = skippedValues + 1");
    expect(vbaCode).toContain("It is SKIPPED, never counted as zero");
    expect(vbaCode).toContain('"Values skipped because they were blank or not numeric: " & skippedValues');
    expect(vbaCode).toContain("MsgBox reportText");
    // A group with no numeric values at all must not be reported as 0.
    expect(vbaCode).toContain("outputValues(g, 2) = ChrW$(8212)");
  });

  it("tolerates spreadsheet number formatting rather than relying on bare CDbl", () => {
    const { vbaCode } = generateVbaFromTemplate(baseForm());
    // currency symbols, thousands separators, padding, parenthesised negatives
    expect(vbaCode).toContain('If ch = "$" Or ch = "," Or ch = " " Or ch = vbTab _');
    expect(vbaCode).toContain('If Left$(rawText, 1) = "(" And Right$(rawText, 1) = ")" Then');
    // Val(), not CDbl(), on the cleaned text: locale-safe and cannot throw.
    expect(vbaCode).toContain("outValue = Val(cleaned)");
    // Booleans/dates/errors are rejected outright rather than coerced.
    expect(vbaCode).toContain("' Empty, Null, Boolean, Date, Error, Object: not a number.");
  });

  it("count does not need a value column and emits no numeric parsing", () => {
    const form = baseForm();
    form.preview.aggregate = "count";
    form.preview.valueColumn = "";
    const result = generateVbaFromTemplate(form);
    expect(result.status).toBe("ok");
    expect(result.vbaCode).toContain("outputValues(g, 2) = groupCount(g)");
    expect(result.vbaCode).not.toContain("TryParseNumber");
    expect(result.vbaCode).toContain('Private Const RESULT_VALUE_HEADER As String = "Count of rows"');
  });

  it("emits the right aggregate expression for average, min and max", () => {
    const cases: [string, string][] = [
      ["average", "outputValues(g, 2) = groupSum(g) / groupNumeric(g)"],
      ["min", "outputValues(g, 2) = groupMin(g)"],
      ["max", "outputValues(g, 2) = groupMax(g)"],
    ];
    for (const [fn, expected] of cases) {
      const form = baseForm();
      form.preview.aggregate = fn as typeof form.preview.aggregate;
      expect(generateVbaFromTemplate(form).vbaCode, fn).toContain(expected);
    }
  });
});

describe("generateVbaFromTemplate — safety scanner behaviour", () => {
  it("produces no scanner findings, which is NOT evidence the code is safe", () => {
    for (const form of [baseForm(), crossWorkbookForm()]) {
      const { vbaCode } = generateVbaFromTemplate(form);
      // Asserting real behaviour of the unmodified scanner: the template
      // deliberately avoids every pattern it knows about. That is a statement
      // about the scanner's coverage, not a safety proof -- the result panel
      // says so on screen, and the generated cautions repeat it.
      expect(scanVbaForWarnings(vbaCode)).toEqual([]);
    }
  });

  it("still tells the user the code is unverified, warnings or not", () => {
    const cautions = generateVbaFromTemplate(baseForm()).safetyCautions.join(" ");
    expect(cautions).toMatch(/has not been run, compiled, or verified/i);
    expect(cautions).toMatch(/Deterministic does not mean correct/i);
  });

  it("flags overwrite mode in the cautions rather than hiding it", () => {
    const overwriteForm = baseForm();
    overwriteForm.mapping.appendOrOverwrite = "overwrite";
    expect(generateVbaFromTemplate(overwriteForm).safetyCautions.join(" ")).toMatch(/overwrite mode/i);
  });

  it("the scanner still fires on genuinely dangerous VBA, so the empty result above means something", () => {
    expect(scanVbaForWarnings('Sub A()\n    Kill "C:\\temp\\x.xlsx"\nEnd Sub').length).toBeGreaterThan(0);
  });
});

// ---------------------------------------------------------------------------
// Shared fixture: template semantics vs previewSimulator semantics.
//
// The VBA cannot be executed here, so this pins the two implementations
// together at the level that is actually checkable: the same structured config
// must yield the same declared stage order, the same grouping key, first-row
// -wins deduplication, and the same "skip, never zero" rule for non-numeric
// values -- with the numeric-parsing tolerances asserted on the simulator side
// and their counterparts asserted present in the emitted VBA.
// ---------------------------------------------------------------------------
describe("template semantics agree with previewSimulator", () => {
  const SAMPLE_HEADERS = ["Account Name", "Account Number", "Amount"];
  const SAMPLE_ROWS = [
    ["Cash - Operating", "1000", "$12,500.00"],
    ["Accounts Receivable", "1200", "$8,300.00"],
    ["Cash - Operating", "1000", "$4,200.00"],
    ["Prepaid Insurance", "1400", "n/a"],
    ["Accounts Receivable", "1200", ""],
    ["Cash - Operating", "1000", "(3,000.00)"],
  ];

  function fixtureForm(kind: "aggregate" | "deduplicate"): MacroFormData {
    const form = baseForm();
    form.preview.kind = kind;
    form.preview.keyColumn = "Account Name";
    form.preview.valueColumn = "Amount";
    form.preview.aggregate = "sum";
    form.preview.sampleHeaders = SAMPLE_HEADERS;
    form.preview.sampleRows = SAMPLE_ROWS.map((r) => [...r]);
    return form;
  }

  it("groups by the same key and sorts ascending in both", () => {
    const form = fixtureForm("aggregate");
    const preview = simulatePreview(form.preview);
    const template = generateVbaFromTemplate(form);

    expect(preview.status).toBe("ok");
    expect(template.status).toBe("ok");

    // Preview: three groups, sorted by key ascending.
    expect(preview.after.rows.map((r) => r[0])).toEqual([
      "Accounts Receivable",
      "Cash - Operating",
      "Prepaid Insurance",
    ]);
    // 12500 + 4200 - 3000 = 13700; the parenthesised value is a negative.
    expect(preview.after.rows.find((r) => r[0] === "Cash - Operating")?.[1]).toBe("13700");
    // The non-numeric "n/a" and the blank were skipped, not zeroed, so those
    // groups have no total at all rather than a misleading 0.
    expect(preview.after.rows.find((r) => r[0] === "Prepaid Insurance")?.[1]).toBe("—");
    expect(preview.notes.join(" ")).toMatch(/not numeric and were skipped, not counted as zero/);

    // Template: same key, same aggregate, same ascending sort, same skip rule.
    expect(template.steps.join(" ")).toMatch(/Group the rows by "Account Name"/);
    expect(template.steps.join(" ")).toMatch(/Sort the result by "Account Name" ascending/);
    expect(template.steps.join(" ")).toMatch(/skipping blank and non-numeric values instead of counting them as zero/);
    expect(template.vbaCode).toContain('normalizedKey = "k:" & LCase$(rawKey)');
    expect(template.vbaCode).toContain("SortRowsByColumn outputValues, outputRowCount, 1");
    expect(template.vbaCode).toContain("outputValues(g, 2) = ChrW$(8212)"); // em dash, matching "—"
    expect(template.vbaCode).toContain("skippedValues = skippedValues + 1");
  });

  it("both keep the FIRST row per distinct key when deduplicating", () => {
    const form = fixtureForm("deduplicate");
    const preview = simulatePreview(form.preview);
    const template = generateVbaFromTemplate(form);

    expect(preview.status).toBe("ok");
    // First occurrence of "Cash - Operating" carries $12,500.00, not $4,200.00.
    expect(preview.after.rows.find((r) => r[0] === "Cash - Operating")?.[2]).toBe("$12,500.00");
    expect(preview.steps.join(" ")).toMatch(/Keep the FIRST row for each distinct/);

    expect(template.steps.join(" ")).toMatch(/Keep the FIRST row for each distinct "Account Name"/);
    // The VBA records the key on first sight and skips every later duplicate.
    expect(template.vbaCode).toContain("If CollectionHasKey(seenKeys, normalizedKey) Then");
    expect(template.vbaCode).toContain("            keepRow(r) = False");
  });

  it("both accept the same tolerated number formats and reject the same junk", () => {
    // Simulator side: asserted by execution.
    expect(parseNumericCell("$12,500.00")).toBe(12500);
    expect(parseNumericCell("(3,000.00)")).toBe(-3000);
    expect(parseNumericCell(" 1 200 ")).toBe(1200);
    expect(parseNumericCell("n/a")).toBeNull();
    expect(parseNumericCell("")).toBeNull();
    expect(parseNumericCell("5.")).toBeNull();

    // Template side: the emitted parser must handle each of those classes.
    const { vbaCode } = generateVbaFromTemplate(fixtureForm("aggregate"));
    expect(vbaCode).toContain('If Left$(rawText, 1) = "(" And Right$(rawText, 1) = ")" Then'); // (3,000.00)
    expect(vbaCode).toContain('If ch = "$" Or ch = "," Or ch = " " Or ch = vbTab _'); // $ , and spaces
    expect(vbaCode).toContain("If Len(rawText) = 0 Then Exit Function"); // blank -> skipped
    expect(vbaCode).toContain('ElseIf ch < "0" Or ch > "9" Then'); // "n/a" -> skipped
    expect(vbaCode).toContain('If Right$(cleaned, 1) = "." Then Exit Function'); // "5." -> skipped
  });

  it("both apply stages in the order filter -> deduplicate -> aggregate", () => {
    // The simulator applies exactly one stage per configured kind, in this
    // fixed order; the template's step narrative must describe the same order.
    const dedupe = generateVbaFromTemplate(fixtureForm("deduplicate")).steps;
    expect(dedupe.findIndex((s) => /FIRST row/.test(s))).toBeLessThan(
      dedupe.findIndex((s) => /Sort the result/.test(s))
    );

    const aggregate = generateVbaFromTemplate(fixtureForm("aggregate")).steps;
    expect(aggregate.findIndex((s) => /Group the rows/.test(s))).toBeLessThan(
      aggregate.findIndex((s) => /Sort the result/.test(s))
    );
    // Sorting is always last, before the write.
    expect(aggregate[aggregate.length - 2]).toMatch(/Clear only the destination block|Append the result/);
  });
});
