import {
  AggregateFn,
  FilterOperator,
  MacroFormData,
  Maybe,
  PreviewConfig,
} from "./types";
import { parseSampleData } from "./previewSimulator";
import { validateVbaMacroName } from "./vbaValidation";

// ---------------------------------------------------------------------------
// Deterministic VBA template generator. No AI, no network, no API key.
//
// SEMANTIC COUNTERPART: `src/lib/previewSimulator.ts`.
//
// The whole value of this module is that the emitted VBA and the on-screen
// before/after preview are built from the SAME structured `PreviewConfig`, so
// they cannot disagree. That only holds while the two implementations stay in
// step. If you change a rule in one file — stage order, first-row-wins
// deduplication, non-numeric skipping, numeric parsing tolerance, sort order —
// you MUST change it in the other, and update the shared-fixture test in
// `vbaTemplateGenerator.test.ts` that pins them together.
//
// The shared rules, in one place:
//   * Stage order is filter -> deduplicate -> aggregate, then sort.
//   * Deduplicate keeps the FIRST row per distinct key.
//   * Aggregate SKIPS non-numeric and blank values. It never coerces them to
//     zero, and it reports how many it skipped.
//   * Numeric parsing tolerates currency symbols, thousands separators and
//     parenthesised negatives.
//   * Results are sorted by key ascending, case-insensitively.
//
// What this module deliberately does NOT do: interpret any free-text rule the
// user typed. Those fields are listed verbatim in `unimplementedRules` and
// echoed into the emitted code as a marked comment block, because clean-looking
// code that silently ignores "skip rows where Status = Cancelled" is worse than
// code that admits it did not read that sentence.
//
// Pure functions only: no DOM, no network, no clock.
// ---------------------------------------------------------------------------

export interface UnimplementedRule {
  field: string;
  text: string;
}

export interface TemplateVbaResult {
  status: "ok" | "unsupported";
  /** Why it cannot build, when status is "unsupported". */
  message?: string;
  vbaCode: string;
  summary: string;
  /** What the emitted code does, in order. */
  steps: string[];
  /** Free-text rules the template did NOT encode. Must be surfaced to the user. */
  unimplementedRules: UnimplementedRule[];
  assumptions: string[];
  installInstructions: string;
  testPlan: string[];
  safetyCautions: string[];
  platformLimitations: string[];
}

const AGGREGATE_LABELS: Record<AggregateFn, string> = {
  sum: "Sum",
  count: "Count",
  average: "Average",
  min: "Min",
  max: "Max",
};

const FILTER_LABELS: Record<FilterOperator, string> = {
  equals: "equals",
  "not-equals": "does not equal",
  contains: "contains",
  "greater-than": "is greater than",
  "less-than": "is less than",
  "is-blank": "is blank",
  "is-not-blank": "is not blank",
};

function unsupported(message: string): TemplateVbaResult {
  return {
    status: "unsupported",
    message,
    // Never hand back half-built code alongside a refusal.
    vbaCode: "",
    summary: "",
    steps: [],
    unimplementedRules: [],
    assumptions: [],
    installInstructions: "",
    testPlan: [],
    safetyCautions: [],
    platformLimitations: [],
  };
}

/** Escapes a value for use inside a VBA double-quoted string literal. */
function vbaString(raw: string): string {
  return `"${(raw ?? "").replace(/"/g, '""')}"`;
}

/** Strips characters that would break out of a single-line VBA comment. */
function vbaComment(raw: string): string {
  return (raw ?? "").replace(/\r?\n/g, " ").trim();
}

function filledMaybe(m: Maybe | undefined): string | null {
  if (!m || m.notApplicable) return null;
  const value = (m.value ?? "").trim();
  return value.length > 0 ? value : null;
}

/**
 * Lists every free-text rule field the user filled in.
 *
 * Deliberately dumb: it does NOT try to work out whether a sentence happens to
 * be covered by the structured picker. Guessing that would be unreliable, and a
 * wrong "already handled" is far more dangerous than a redundant reminder.
 * Fields marked "Not applicable" or left blank are omitted.
 */
export function collectUnimplementedRules(form: MacroFormData): UnimplementedRule[] {
  const candidates: [string, Maybe | undefined][] = [
    ["Filters", form.mapping?.filters],
    ["Transformation/calculation rules", form.mapping?.transformationRules],
    ["Sort order", form.mapping?.sortOrder],
    ["Duplicate handling", form.mapping?.duplicateHandling],
    ["Blank/error handling", form.mapping?.blankOrErrorHandling],
    ["Optional constraints", form.run?.constraints],
    ["Anything the macro must not change", form.run?.mustNotChange],
  ];

  const rules: UnimplementedRule[] = [];
  for (const [field, maybe] of candidates) {
    const text = filledMaybe(maybe);
    if (text !== null) rules.push({ field, text });
  }
  return rules;
}

// ---------------------------------------------------------------------------
// VBA emission
// ---------------------------------------------------------------------------

interface Plan {
  macroName: string;
  sourceWorkbook: string;
  sourceWorksheet: string;
  sourceRangeOrTable: string;
  sameWorkbook: boolean;
  destinationWorkbook: string;
  destinationWorksheet: string;
  destinationRangeOrTable: string;
  overwrite: boolean;
  kind: Exclude<PreviewConfig["kind"], "not-configured">;
  keyColumn: string;
  valueColumn: string;
  aggregate: AggregateFn;
  filterOperator: FilterOperator;
  filterValue: string;
  /** Column headers as parsed from step 2, in source order. */
  headers: string[];
  trigger: "manual" | "button";
}

/** Header written above the aggregated value column. */
/**
 * The headings the macro writes, in the order it writes them. An aggregate
 * collapses to key + aggregated value; every other operation keeps the source
 * headings. Used both for the destination-table header match and to tell the
 * user which headings their table needs.
 */
function outputHeaderNames(plan: Plan): string[] {
  if (plan.kind === "aggregate") return [plan.keyColumn, resultValueHeader(plan)];
  return plan.headers;
}

function resultValueHeader(plan: Plan): string {
  if (plan.aggregate === "count") return "Count of rows";
  return `${AGGREGATE_LABELS[plan.aggregate]} of ${plan.valueColumn}`;
}

function needsValueColumn(plan: Plan): boolean {
  return plan.kind === "aggregate" && plan.aggregate !== "count";
}

function needsNumberParser(plan: Plan): boolean {
  if (needsValueColumn(plan)) return true;
  return (
    plan.kind === "filter" &&
    (plan.filterOperator === "greater-than" || plan.filterOperator === "less-than")
  );
}

function needsSort(plan: Plan): boolean {
  return plan.kind === "deduplicate" || plan.kind === "aggregate";
}

function needsKeyLookup(plan: Plan): boolean {
  return plan.kind !== "copy";
}

function buildHeaderBlock(plan: Plan, rules: UnimplementedRule[], steps: string[]): string[] {
  const lines: string[] = [];
  lines.push("'=====================================================================");
  lines.push(`'  ${plan.macroName}`);
  lines.push("'  Generated by Excel Macro Builder from a DETERMINISTIC TEMPLATE.");
  lines.push("'  No AI model wrote this code. It was assembled from the structured");
  lines.push("'  operation picked in step 2, so it matches the on-screen preview.");
  lines.push("'");
  lines.push("'  What it does, in order:");
  for (const step of steps) {
    lines.push(`'    - ${vbaComment(step)}`);
  }
  lines.push("'");
  lines.push("'  ------------------------------------------------------------------");
  lines.push("'  RULES FROM YOUR DESCRIPTION THAT WERE **NOT** IMPLEMENTED");
  lines.push("'  ------------------------------------------------------------------");
  if (rules.length === 0) {
    lines.push("'  You did not leave any free-text rules filled in, so there is");
    lines.push("'  nothing outstanding from those fields. This template still only");
    lines.push("'  implements the structured operation above -- it never reads or");
    lines.push("'  interprets free text of any kind.");
  } else {
    lines.push("'  The template does not read free text. Review each rule below and");
    lines.push("'  add it to the code yourself before relying on this macro:");
    for (const rule of rules) {
      lines.push(`'    - ${vbaComment(rule.field)}: ${vbaComment(rule.text)}`);
    }
  }
  lines.push("'");
  lines.push("'  This code has NOT been run, compiled, or verified by anything.");
  lines.push("'  Read every line, then test it on a COPY of your workbook.");
  lines.push("'=====================================================================");
  return lines;
}

function buildConfigBlock(plan: Plan): string[] {
  const lines: string[] = [];
  lines.push("' --------------------------------------------------------------------");
  lines.push("' CONFIGURATION -- edit these, and nothing below, to retarget the macro.");
  lines.push("' --------------------------------------------------------------------");
  lines.push(`Private Const SOURCE_WORKBOOK_NAME As String = ${vbaString(plan.sourceWorkbook)}`);
  lines.push(`Private Const SOURCE_SHEET_NAME As String = ${vbaString(plan.sourceWorksheet)}`);
  lines.push(`Private Const SOURCE_RANGE_OR_TABLE As String = ${vbaString(plan.sourceRangeOrTable)}`);
  if (!plan.sameWorkbook) {
    lines.push("' Full path is best here. A bare file name only works if the workbook");
    lines.push("' is already open, or sits in Excel's current default folder.");
    lines.push(`Private Const DEST_WORKBOOK_PATH As String = ${vbaString(plan.destinationWorkbook)}`);
  }
  lines.push(`Private Const DEST_SHEET_NAME As String = ${vbaString(plan.destinationWorksheet)}`);
  lines.push(`Private Const DEST_RANGE_OR_TABLE As String = ${vbaString(plan.destinationRangeOrTable)}`);
  if (needsKeyLookup(plan)) {
    lines.push(`Private Const KEY_HEADER As String = ${vbaString(plan.keyColumn)}`);
  }
  if (needsValueColumn(plan)) {
    lines.push(`Private Const VALUE_HEADER As String = ${vbaString(plan.valueColumn)}`);
  }
  if (plan.kind === "aggregate") {
    lines.push(`Private Const RESULT_VALUE_HEADER As String = ${vbaString(resultValueHeader(plan))}`);
  }
  if (plan.kind === "filter" && plan.filterOperator !== "is-blank" && plan.filterOperator !== "is-not-blank") {
    lines.push(`Private Const FILTER_VALUE As String = ${vbaString(plan.filterValue)}`);
  }
  lines.push(`Private Const MACRO_LABEL As String = ${vbaString(plan.macroName)}`);
  lines.push("Private Const ERR_BASE As Long = vbObjectError + 3000");
  lines.push("' --------------------------------------------------------------------");
  lines.push("' END CONFIGURATION");
  lines.push("' --------------------------------------------------------------------");
  return lines;
}

function buildDimBlock(plan: Plan): string[] {
  const lines: string[] = [];
  lines.push("    Dim prevScreenUpdating As Boolean");
  lines.push("    Dim prevEnableEvents As Boolean");
  lines.push("    Dim prevDisplayAlerts As Boolean");
  lines.push("    Dim prevCalculation As Long");
  lines.push("    Dim stateSaved As Boolean");
  lines.push("    Dim failed As Boolean");
  lines.push("    Dim failMessage As String");
  lines.push("    Dim reportText As String");
  lines.push("");
  lines.push("    Dim sourceBook As Workbook");
  lines.push("    Dim sourceSheet As Worksheet");
  lines.push("    Dim sourceRange As Range");
  lines.push("    Dim headerRow As Range");
  lines.push("    Dim dataBody As Range");
  lines.push("    Dim sourceValues As Variant");
  lines.push("");
  lines.push("    Dim destBook As Workbook");
  lines.push("    Dim destSheet As Worksheet");
  lines.push("    Dim destTable As ListObject");
  lines.push("    Dim destAnchor As Range");
  lines.push("");
  lines.push("    Dim rowCount As Long");
  lines.push("    Dim columnCount As Long");
  lines.push("    Dim outputHeaders() As Variant");
  lines.push("    Dim outputValues() As Variant");
  lines.push("    Dim outputColumns As Long");
  lines.push("    Dim outputRowCount As Long");
  lines.push("    Dim r As Long");
  lines.push("    Dim c As Long");
  lines.push("    Dim destColumnMap() As Long");
  lines.push("    Dim columnBuffer() As Variant");

  if (needsKeyLookup(plan)) lines.push("    Dim keyColumn As Long");
  if (needsValueColumn(plan)) lines.push("    Dim valueColumn As Long");

  if (plan.kind === "filter" || plan.kind === "deduplicate") {
    lines.push("    Dim keepRow() As Boolean");
    lines.push("    Dim keptCount As Long");
    lines.push("    Dim outIndex As Long");
  }
  if (plan.kind === "filter" && needsNumberParser(plan)) {
    lines.push("    Dim leftNumber As Double");
    lines.push("    Dim rightNumber As Double");
  }
  if (plan.kind === "deduplicate") {
    lines.push("    Dim seenKeys As Collection");
    lines.push("    Dim normalizedKey As String");
  }
  if (plan.kind === "aggregate") {
    lines.push("    Dim groupIndex As Collection");
    lines.push("    Dim groupLabels() As String");
    lines.push("    Dim groupCount() As Long");
    lines.push("    Dim groupTotal As Long");
    lines.push("    Dim g As Long");
    lines.push("    Dim rawKey As String");
    lines.push("    Dim normalizedKey As String");
    if (needsValueColumn(plan)) {
      lines.push("    Dim groupSum() As Double");
      lines.push("    Dim groupMin() As Double");
      lines.push("    Dim groupMax() As Double");
      lines.push("    Dim groupNumeric() As Long");
      lines.push("    Dim parsedNumber As Double");
      lines.push("    Dim skippedValues As Long");
    }
  }
  if (plan.overwrite) {
    lines.push("    Dim clearLastRow As Long");
  } else {
    lines.push("    Dim lastUsedRow As Long");
    lines.push("    Dim writeStart As Range");
    lines.push("    Dim newRow As ListRow");
  }
  return lines;
}

function buildReadSourceBlock(plan: Plan): string[] {
  const lines: string[] = [];
  lines.push("    ' --- Locate the source data ---------------------------------------");
  lines.push("    Set sourceBook = FindOpenWorkbook(SOURCE_WORKBOOK_NAME)");
  lines.push("    If sourceBook Is Nothing Then");
  lines.push(
    '        Err.Raise ERR_BASE + 1, MACRO_LABEL, "Source workbook """ & SOURCE_WORKBOOK_NAME & """ is not open in Excel. Open it and run this macro again."'
  );
  lines.push("    End If");
  lines.push("");
  lines.push("    Set sourceSheet = FindWorksheet(sourceBook, SOURCE_SHEET_NAME)");
  lines.push("    If sourceSheet Is Nothing Then");
  lines.push(
    '        Err.Raise ERR_BASE + 2, MACRO_LABEL, "Worksheet """ & SOURCE_SHEET_NAME & """ was not found in """ & SOURCE_WORKBOOK_NAME & """."'
  );
  lines.push("    End If");
  lines.push("");
  lines.push("    Set sourceRange = ResolveDataRange(sourceSheet, SOURCE_RANGE_OR_TABLE)");
  lines.push("    If sourceRange Is Nothing Then");
  lines.push(
    '        Err.Raise ERR_BASE + 3, MACRO_LABEL, "No table or range named """ & SOURCE_RANGE_OR_TABLE & """ was found on worksheet """ & SOURCE_SHEET_NAME & """."'
  );
  lines.push("    End If");
  lines.push("");
  lines.push("    ' The first row of the source range is treated as the header row.");
  lines.push("    If sourceRange.Rows.Count < 2 Then");
  lines.push(
    '        Err.Raise ERR_BASE + 4, MACRO_LABEL, "The source range """ & SOURCE_RANGE_OR_TABLE & """ has no data rows underneath its header row."'
  );
  lines.push("    End If");
  lines.push("");
  lines.push("    Set headerRow = sourceRange.Rows(1)");
  lines.push("    Set dataBody = sourceRange.Offset(1, 0).Resize(sourceRange.Rows.Count - 1, sourceRange.Columns.Count)");
  lines.push("    sourceValues = ToArray2D(dataBody)");
  lines.push("    rowCount = UBound(sourceValues, 1)");
  lines.push("    columnCount = UBound(sourceValues, 2)");
  lines.push("");

  if (needsKeyLookup(plan)) {
    lines.push("    ' Columns are found BY HEADER NAME, never by a hard-coded index.");
    lines.push("    keyColumn = ColumnIndexByHeader(headerRow, KEY_HEADER)");
    lines.push("    If keyColumn = 0 Then");
    lines.push(
      '        Err.Raise ERR_BASE + 5, MACRO_LABEL, "Header """ & KEY_HEADER & """ was not found in the source header row of """ & SOURCE_RANGE_OR_TABLE & """."'
    );
    lines.push("    End If");
  }
  if (needsValueColumn(plan)) {
    lines.push("    valueColumn = ColumnIndexByHeader(headerRow, VALUE_HEADER)");
    lines.push("    If valueColumn = 0 Then");
    lines.push(
      '        Err.Raise ERR_BASE + 6, MACRO_LABEL, "Header """ & VALUE_HEADER & """ was not found in the source header row of """ & SOURCE_RANGE_OR_TABLE & """."'
    );
    lines.push("    End If");
  }
  return lines;
}

function filterConditionLines(plan: Plan): string[] {
  const cell = "CellText(sourceValues(r, keyColumn))";
  switch (plan.filterOperator) {
    case "is-blank":
      return [`        keepRow(r) = (Len(${cell}) = 0)`];
    case "is-not-blank":
      return [`        keepRow(r) = (Len(${cell}) > 0)`];
    case "equals":
      return [`        keepRow(r) = (StrComp(${cell}, FILTER_VALUE, vbTextCompare) = 0)`];
    case "not-equals":
      return [`        keepRow(r) = (StrComp(${cell}, FILTER_VALUE, vbTextCompare) <> 0)`];
    case "contains":
      return [`        keepRow(r) = (InStr(1, ${cell}, FILTER_VALUE, vbTextCompare) > 0)`];
    case "greater-than":
    case "less-than": {
      const numeric = plan.filterOperator === "greater-than" ? ">" : "<";
      return [
        "        ' Compare as numbers when both sides parse, otherwise fall back to",
        "        ' a case-insensitive text comparison (same rule as the preview).",
        `        If TryParseNumber(sourceValues(r, keyColumn), leftNumber) And TryParseNumber(FILTER_VALUE, rightNumber) Then`,
        `            keepRow(r) = (leftNumber ${numeric} rightNumber)`,
        "        Else",
        `            keepRow(r) = (StrComp(${cell}, FILTER_VALUE, vbTextCompare) ${numeric} 0)`,
        "        End If",
      ];
    }
    default:
      return ["        keepRow(r) = True"];
  }
}

function buildTransformBlock(plan: Plan): string[] {
  const lines: string[] = [];

  if (plan.kind === "copy") {
    lines.push("    ' --- Copy every source row through unchanged -----------------------");
    lines.push("    outputColumns = columnCount");
    lines.push("    outputRowCount = rowCount");
    lines.push("    ReDim outputHeaders(1 To 1, 1 To outputColumns)");
    lines.push("    For c = 1 To outputColumns");
    lines.push("        outputHeaders(1, c) = headerRow.Cells(1, c).Value");
    lines.push("    Next c");
    lines.push("    ReDim outputValues(1 To outputRowCount, 1 To outputColumns)");
    lines.push("    For r = 1 To rowCount");
    lines.push("        For c = 1 To columnCount");
    lines.push("            outputValues(r, c) = sourceValues(r, c)");
    lines.push("        Next c");
    lines.push("    Next r");
    return lines;
  }

  if (plan.kind === "filter") {
    lines.push("    ' --- Stage 1 of 1: filter ------------------------------------------");
    lines.push("    outputColumns = columnCount");
    lines.push("    ReDim outputHeaders(1 To 1, 1 To outputColumns)");
    lines.push("    For c = 1 To outputColumns");
    lines.push("        outputHeaders(1, c) = headerRow.Cells(1, c).Value");
    lines.push("    Next c");
    lines.push("");
    lines.push("    ReDim keepRow(1 To rowCount)");
    lines.push("    keptCount = 0");
    lines.push("    For r = 1 To rowCount");
    lines.push(...filterConditionLines(plan));
    lines.push("        If keepRow(r) Then keptCount = keptCount + 1");
    lines.push("    Next r");
    lines.push("");
    lines.push("    outputRowCount = keptCount");
    lines.push("    If outputRowCount > 0 Then");
    lines.push("        ReDim outputValues(1 To outputRowCount, 1 To outputColumns)");
    lines.push("        outIndex = 0");
    lines.push("        For r = 1 To rowCount");
    lines.push("            If keepRow(r) Then");
    lines.push("                outIndex = outIndex + 1");
    lines.push("                For c = 1 To columnCount");
    lines.push("                    outputValues(outIndex, c) = sourceValues(r, c)");
    lines.push("                Next c");
    lines.push("            End If");
    lines.push("        Next r");
    lines.push("    End If");
    return lines;
  }

  if (plan.kind === "deduplicate") {
    lines.push("    ' --- Stage 1 of 2: deduplicate -------------------------------------");
    lines.push("    ' Keeps the FIRST row seen for each distinct key, matching the");
    lines.push("    ' preview. Keys are compared trimmed and case-insensitively.");
    lines.push("    outputColumns = columnCount");
    lines.push("    ReDim outputHeaders(1 To 1, 1 To outputColumns)");
    lines.push("    For c = 1 To outputColumns");
    lines.push("        outputHeaders(1, c) = headerRow.Cells(1, c).Value");
    lines.push("    Next c");
    lines.push("");
    lines.push("    Set seenKeys = New Collection");
    lines.push("    ReDim keepRow(1 To rowCount)");
    lines.push("    keptCount = 0");
    lines.push("    For r = 1 To rowCount");
    lines.push('        normalizedKey = "k:" & LCase$(CellText(sourceValues(r, keyColumn)))');
    lines.push("        If CollectionHasKey(seenKeys, normalizedKey) Then");
    lines.push("            keepRow(r) = False");
    lines.push("        Else");
    lines.push("            seenKeys.Add True, normalizedKey");
    lines.push("            keepRow(r) = True");
    lines.push("            keptCount = keptCount + 1");
    lines.push("        End If");
    lines.push("    Next r");
    lines.push("");
    lines.push("    outputRowCount = keptCount");
    lines.push("    If outputRowCount > 0 Then");
    lines.push("        ReDim outputValues(1 To outputRowCount, 1 To outputColumns)");
    lines.push("        outIndex = 0");
    lines.push("        For r = 1 To rowCount");
    lines.push("            If keepRow(r) Then");
    lines.push("                outIndex = outIndex + 1");
    lines.push("                For c = 1 To columnCount");
    lines.push("                    outputValues(outIndex, c) = sourceValues(r, c)");
    lines.push("                Next c");
    lines.push("            End If");
    lines.push("        Next r");
    lines.push("");
    lines.push("        ' --- Stage 2 of 2: sort by key ascending ------------------------");
    lines.push("        SortRowsByColumn outputValues, outputRowCount, keyColumn");
    lines.push("    End If");
    return lines;
  }

  // --- aggregate -----------------------------------------------------------
  const withValue = needsValueColumn(plan);
  lines.push("    ' --- Stage 1 of 2: group and aggregate -----------------------------");
  lines.push("    outputColumns = 2");
  lines.push("    ReDim outputHeaders(1 To 1, 1 To 2)");
  lines.push("    outputHeaders(1, 1) = KEY_HEADER");
  lines.push("    outputHeaders(1, 2) = RESULT_VALUE_HEADER");
  lines.push("");
  lines.push("    Set groupIndex = New Collection");
  lines.push("    ReDim groupLabels(1 To rowCount)");
  lines.push("    ReDim groupCount(1 To rowCount)");
  if (withValue) {
    lines.push("    ReDim groupSum(1 To rowCount)");
    lines.push("    ReDim groupMin(1 To rowCount)");
    lines.push("    ReDim groupMax(1 To rowCount)");
    lines.push("    ReDim groupNumeric(1 To rowCount)");
    lines.push("    skippedValues = 0");
  }
  lines.push("    groupTotal = 0");
  lines.push("");
  lines.push("    For r = 1 To rowCount");
  lines.push("        rawKey = CellText(sourceValues(r, keyColumn))");
  lines.push('        normalizedKey = "k:" & LCase$(rawKey)');
  lines.push("        If CollectionHasKey(groupIndex, normalizedKey) Then");
  lines.push("            g = groupIndex.Item(normalizedKey)");
  lines.push("        Else");
  lines.push("            groupTotal = groupTotal + 1");
  lines.push("            g = groupTotal");
  lines.push("            groupIndex.Add g, normalizedKey");
  lines.push("            groupLabels(g) = rawKey");
  lines.push("            groupCount(g) = 0");
  if (withValue) {
    lines.push("            groupNumeric(g) = 0");
    lines.push("            groupSum(g) = 0");
  }
  lines.push("        End If");
  lines.push("        groupCount(g) = groupCount(g) + 1");
  if (withValue) {
    lines.push("");
    lines.push("        If TryParseNumber(sourceValues(r, valueColumn), parsedNumber) Then");
    lines.push("            groupNumeric(g) = groupNumeric(g) + 1");
    lines.push("            groupSum(g) = groupSum(g) + parsedNumber");
    lines.push("            If groupNumeric(g) = 1 Then");
    lines.push("                groupMin(g) = parsedNumber");
    lines.push("                groupMax(g) = parsedNumber");
    lines.push("            Else");
    lines.push("                If parsedNumber < groupMin(g) Then groupMin(g) = parsedNumber");
    lines.push("                If parsedNumber > groupMax(g) Then groupMax(g) = parsedNumber");
    lines.push("            End If");
    lines.push("        Else");
    lines.push("            ' Blank or non-numeric. It is SKIPPED, never counted as zero,");
    lines.push("            ' and the total skipped is reported when the macro finishes.");
    lines.push("            skippedValues = skippedValues + 1");
    lines.push("        End If");
  }
  lines.push("    Next r");
  lines.push("");
  lines.push("    outputRowCount = groupTotal");
  lines.push("    If outputRowCount > 0 Then");
  lines.push("        ReDim outputValues(1 To outputRowCount, 1 To 2)");
  lines.push("        For g = 1 To groupTotal");
  lines.push("            outputValues(g, 1) = groupLabels(g)");
  if (!withValue) {
    lines.push("            outputValues(g, 2) = groupCount(g)");
  } else {
    lines.push("            If groupNumeric(g) = 0 Then");
    lines.push("                ' No numeric value at all in this group, so there is no");
    lines.push("                ' honest total to report. An em dash is written, not a 0.");
    lines.push("                outputValues(g, 2) = ChrW$(8212)");
    lines.push("            Else");
    switch (plan.aggregate) {
      case "sum":
        lines.push("                outputValues(g, 2) = groupSum(g)");
        break;
      case "average":
        lines.push("                outputValues(g, 2) = groupSum(g) / groupNumeric(g)");
        break;
      case "min":
        lines.push("                outputValues(g, 2) = groupMin(g)");
        break;
      case "max":
        lines.push("                outputValues(g, 2) = groupMax(g)");
        break;
      default:
        lines.push("                outputValues(g, 2) = groupSum(g)");
        break;
    }
    lines.push("            End If");
  }
  lines.push("        Next g");
  lines.push("");
  lines.push("        ' --- Stage 2 of 2: sort by key ascending ------------------------");
  lines.push("        SortRowsByColumn outputValues, outputRowCount, 1");
  lines.push("    End If");
  return lines;
}

function buildDestinationBlock(plan: Plan): string[] {
  const lines: string[] = [];
  lines.push("    ' --- Locate the destination ----------------------------------------");
  if (plan.sameWorkbook) {
    lines.push("    ' Source and destination are the same workbook.");
    lines.push("    Set destBook = sourceBook");
  } else {
    lines.push("    ' Reuse the destination workbook if it is already open, so the macro");
    lines.push("    ' never opens a second copy of it.");
    lines.push("    Set destBook = FindOpenWorkbook(FileNameOnly(DEST_WORKBOOK_PATH))");
    lines.push("    If destBook Is Nothing Then");
    lines.push("        Set destBook = Application.Workbooks.Open(FileName:=DEST_WORKBOOK_PATH)");
    lines.push("    End If");
    lines.push("    If destBook Is Nothing Then");
    lines.push(
      '        Err.Raise ERR_BASE + 7, MACRO_LABEL, "Destination workbook """ & DEST_WORKBOOK_PATH & """ could not be found or opened. Check the path."'
    );
    lines.push("    End If");
    lines.push("    ' Deliberately NOT saved and NOT closed: it is left open so you can");
    lines.push("    ' look at the result and decide whether to keep it.");
  }
  lines.push("");
  lines.push("    Set destSheet = FindWorksheet(destBook, DEST_SHEET_NAME)");
  lines.push("    If destSheet Is Nothing Then");
  lines.push(
    '        Err.Raise ERR_BASE + 8, MACRO_LABEL, "Worksheet """ & DEST_SHEET_NAME & """ was not found in the destination workbook."'
  );
  lines.push("    End If");
  lines.push("");
  lines.push("    Set destTable = FindListObject(destSheet, DEST_RANGE_OR_TABLE)");
  lines.push("    If destTable Is Nothing Then");
  lines.push("        Set destAnchor = ResolveAnchorCell(destSheet, DEST_RANGE_OR_TABLE)");
  lines.push("        If destAnchor Is Nothing Then");
  lines.push(
    '            Err.Raise ERR_BASE + 9, MACRO_LABEL, """" & DEST_RANGE_OR_TABLE & """ is neither a table on """ & DEST_SHEET_NAME & """ nor a usable cell/range address."'
  );
  lines.push("        End If");
  lines.push("    End If");
  lines.push("");

  if (plan.overwrite) {
    lines.push("    ' --- Write (overwrite) ---------------------------------------------");
    lines.push("    If Not destTable Is Nothing Then");
    lines.push("        MapOutputColumnsToTable destTable, outputHeaders, outputColumns, destColumnMap");
    lines.push("        ' Clear only the table's own data body. The table header, the rest");
    lines.push("        ' of the sheet, and every row outside the table are untouched, and");
    lines.push("        ' no rows are deleted anywhere.");
    lines.push("        If Not destTable.DataBodyRange Is Nothing Then destTable.DataBodyRange.ClearContents");
    lines.push("        destTable.Resize destTable.Range.Resize(outputRowCount + 1, destTable.Range.Columns.Count)");
    lines.push("        If outputRowCount > 0 Then");
    lines.push("            ' Written column by column into the table column whose HEADER");
    lines.push("            ' matches, never by position, so a reordered or renamed table");
    lines.push("            ' cannot silently put values under the wrong heading.");
    lines.push("            For c = 1 To outputColumns");
    lines.push("                ReDim columnBuffer(1 To outputRowCount, 1 To 1)");
    lines.push("                For r = 1 To outputRowCount");
    lines.push("                    columnBuffer(r, 1) = outputValues(r, c)");
    lines.push("                Next r");
    lines.push("                destTable.DataBodyRange.Cells(1, destColumnMap(c)).Resize(outputRowCount, 1).Value = columnBuffer");
    lines.push("            Next c");
    lines.push("        End If");
    lines.push("    Else");
    lines.push("        ' Clear only the block this macro owns: from the anchor cell down,");
    lines.push("        ' and only across the columns it writes. Nothing outside that");
    lines.push("        ' rectangle is cleared, and no rows are deleted.");
    lines.push("        clearLastRow = destSheet.Cells(destSheet.Rows.Count, destAnchor.Column).End(xlUp).Row");
    lines.push("        If clearLastRow < destAnchor.Row Then clearLastRow = destAnchor.Row");
    lines.push("        destSheet.Range( _");
    lines.push("            destSheet.Cells(destAnchor.Row, destAnchor.Column), _");
    lines.push("            destSheet.Cells(clearLastRow, destAnchor.Column + outputColumns - 1)).ClearContents");
    lines.push("        destAnchor.Resize(1, outputColumns).Value = outputHeaders");
    lines.push("        If outputRowCount > 0 Then");
    lines.push("            destAnchor.Offset(1, 0).Resize(outputRowCount, outputColumns).Value = outputValues");
    lines.push("        End If");
    lines.push("    End If");
  } else {
    lines.push("    ' --- Write (append) -------------------------------------------------");
    lines.push("    ' Nothing already in the destination is cleared, changed or deleted.");
    lines.push("    If Not destTable Is Nothing Then");
    lines.push("        MapOutputColumnsToTable destTable, outputHeaders, outputColumns, destColumnMap");
    lines.push("        ' Each value goes into the table column whose HEADER matches, never");
    lines.push("        ' by position, so a reordered or renamed table cannot silently put");
    lines.push("        ' values under the wrong heading.");
    lines.push("        For r = 1 To outputRowCount");
    lines.push("            Set newRow = destTable.ListRows.Add");
    lines.push("            For c = 1 To outputColumns");
    lines.push("                newRow.Range.Cells(1, destColumnMap(c)).Value = outputValues(r, c)");
    lines.push("            Next c");
    lines.push("        Next r");
    lines.push("    Else");
    lines.push("        lastUsedRow = destSheet.Cells(destSheet.Rows.Count, destAnchor.Column).End(xlUp).Row");
    lines.push("        If lastUsedRow < destAnchor.Row Then lastUsedRow = destAnchor.Row");
    lines.push("        If Len(CellText(destAnchor.Value)) = 0 And lastUsedRow = destAnchor.Row Then");
    lines.push("            ' Empty destination: write the header row once, then the data.");
    lines.push("            destAnchor.Resize(1, outputColumns).Value = outputHeaders");
    lines.push("            Set writeStart = destAnchor.Offset(1, 0)");
    lines.push("        Else");
    lines.push("            Set writeStart = destSheet.Cells(lastUsedRow + 1, destAnchor.Column)");
    lines.push("        End If");
    lines.push("        If outputRowCount > 0 Then");
    lines.push("            writeStart.Resize(outputRowCount, outputColumns).Value = outputValues");
    lines.push("        End If");
    lines.push("    End If");
  }
  return lines;
}

function buildReportBlock(plan: Plan): string[] {
  const lines: string[] = [];
  const modeWord = plan.overwrite ? "overwrite" : "append";
  lines.push("    ' --- Report back ----------------------------------------------------");
  lines.push(
    `    reportText = MACRO_LABEL & " finished." & vbCrLf & _`
  );
  lines.push('        "Source rows read: " & rowCount & vbCrLf & _');
  lines.push(
    `        "Rows written to '" & DEST_SHEET_NAME & "' (${modeWord}): " & outputRowCount & vbCrLf & _`
  );
  if (needsValueColumn(plan)) {
    lines.push(
      '        "Values skipped because they were blank or not numeric: " & skippedValues & _'
    );
    lines.push('        " (skipped values are NEVER counted as zero)." & vbCrLf & _');
  }
  lines.push(
    '        vbCrLf & "This macro was built from a template and has not been verified. Check the result."'
  );
  return lines;
}

function buildHelpers(plan: Plan): string[] {
  const lines: string[] = [];

  lines.push("' =====================================================================");
  lines.push("' Helper functions");
  lines.push("' =====================================================================");
  lines.push("");
  lines.push("' Returns the open workbook with this file name, or Nothing.");
  lines.push("Private Function FindOpenWorkbook(ByVal workbookName As String) As Workbook");
  lines.push("    Dim candidate As Workbook");
  lines.push("    For Each candidate In Application.Workbooks");
  lines.push("        If StrComp(candidate.Name, workbookName, vbTextCompare) = 0 Then");
  lines.push("            Set FindOpenWorkbook = candidate");
  lines.push("            Exit Function");
  lines.push("        End If");
  lines.push("    Next candidate");
  lines.push("End Function");
  lines.push("");

  if (!plan.sameWorkbook) {
    lines.push("' Strips any folder part from a path, for Windows and Mac separators.");
    lines.push("Private Function FileNameOnly(ByVal fullPath As String) As String");
    lines.push("    Dim i As Long");
    lines.push("    Dim ch As String");
    lines.push("    For i = Len(fullPath) To 1 Step -1");
    lines.push("        ch = Mid$(fullPath, i, 1)");
    lines.push('        If ch = "\\" Or ch = "/" Or ch = ":" Then');
    lines.push("            FileNameOnly = Mid$(fullPath, i + 1)");
    lines.push("            Exit Function");
    lines.push("        End If");
    lines.push("    Next i");
    lines.push("    FileNameOnly = fullPath");
    lines.push("End Function");
    lines.push("");
  }

  lines.push("' Returns the named worksheet in a workbook, or Nothing.");
  lines.push("Private Function FindWorksheet(ByVal book As Workbook, ByVal sheetName As String) As Worksheet");
  lines.push("    Dim candidate As Worksheet");
  lines.push("    For Each candidate In book.Worksheets");
  lines.push("        If StrComp(candidate.Name, sheetName, vbTextCompare) = 0 Then");
  lines.push("            Set FindWorksheet = candidate");
  lines.push("            Exit Function");
  lines.push("        End If");
  lines.push("    Next candidate");
  lines.push("End Function");
  lines.push("");

  lines.push("' Returns the named table on a worksheet, or Nothing.");
  lines.push("Private Function FindListObject(ByVal sheet As Worksheet, ByVal tableName As String) As ListObject");
  lines.push("    Dim candidate As ListObject");
  lines.push("    For Each candidate In sheet.ListObjects");
  lines.push("        If StrComp(candidate.Name, tableName, vbTextCompare) = 0 Then");
  lines.push("            Set FindListObject = candidate");
  lines.push("            Exit Function");
  lines.push("        End If");
  lines.push("    Next candidate");
  lines.push("End Function");
  lines.push("");

  lines.push("' Resolves a name that may be either a table or a range address, header");
  lines.push("' row included. Returns Nothing when neither resolves.");
  lines.push("Private Function ResolveDataRange(ByVal sheet As Worksheet, ByVal rangeOrTable As String) As Range");
  lines.push("    Dim foundTable As ListObject");
  lines.push("    Dim resolved As Range");
  lines.push("    Set foundTable = FindListObject(sheet, rangeOrTable)");
  lines.push("    If Not foundTable Is Nothing Then");
  lines.push("        Set ResolveDataRange = foundTable.Range");
  lines.push("        Exit Function");
  lines.push("    End If");
  lines.push("    On Error Resume Next");
  lines.push("    Set resolved = sheet.Range(rangeOrTable)");
  lines.push("    On Error GoTo 0");
  lines.push("    Set ResolveDataRange = resolved");
  lines.push("End Function");
  lines.push("");

  lines.push("' Resolves a range address to its top-left cell. Returns Nothing when the");
  lines.push("' address cannot be resolved.");
  lines.push("Private Function ResolveAnchorCell(ByVal sheet As Worksheet, ByVal rangeAddress As String) As Range");
  lines.push("    Dim resolved As Range");
  lines.push("    On Error Resume Next");
  lines.push("    Set resolved = sheet.Range(rangeAddress)");
  lines.push("    On Error GoTo 0");
  lines.push("    If resolved Is Nothing Then Exit Function");
  lines.push("    Set ResolveAnchorCell = resolved.Cells(1, 1)");
  lines.push("End Function");
  lines.push("");

  lines.push("' Finds a column by its header text. Returns 0 when it is not present,");
  lines.push("' so the caller can raise an error naming the missing header.");
  lines.push("Private Function ColumnIndexByHeader(ByVal headers As Range, ByVal headerName As String) As Long");
  lines.push("    Dim i As Long");
  lines.push("    For i = 1 To headers.Cells.Count");
  lines.push("        If StrComp(CellText(headers.Cells(1, i).Value), headerName, vbTextCompare) = 0 Then");
  lines.push("            ColumnIndexByHeader = i");
  lines.push("            Exit Function");
  lines.push("        End If");
  lines.push("    Next i");
  lines.push("    ColumnIndexByHeader = 0");
  lines.push("End Function");
  lines.push("");

  lines.push("' Works out which column of the destination table each output column");
  lines.push("' belongs in, by matching HEADER TEXT rather than position. Without this,");
  lines.push("' a table whose columns are renamed or reordered would take the values");
  lines.push("' positionally and file them under the wrong heading without complaining.");
  lines.push("' Raises a descriptive error instead of writing anything questionable.");
  lines.push("Private Sub MapOutputColumnsToTable(ByVal targetTable As ListObject, ByRef headers() As Variant, ByVal columnsUsed As Long, ByRef columnMap() As Long)");
  lines.push("    Dim i As Long");
  lines.push("    Dim wantedHeader As String");
  lines.push("");
  lines.push("    If targetTable.HeaderRowRange Is Nothing Then");
  lines.push(
    '        Err.Raise ERR_BASE + 11, MACRO_LABEL, "Destination table """ & targetTable.Name & """ has its header row switched off, so its columns cannot be matched by name. Turn the header row on, or write to a plain range instead."'
  );
  lines.push("    End If");
  lines.push("");
  lines.push("    If targetTable.Range.Columns.Count < columnsUsed Then");
  lines.push(
    '        Err.Raise ERR_BASE + 10, MACRO_LABEL, "Destination table """ & targetTable.Name & """ has " & targetTable.Range.Columns.Count & " column(s), but this macro writes " & columnsUsed & "."'
  );
  lines.push("    End If");
  lines.push("");
  lines.push("    ReDim columnMap(1 To columnsUsed)");
  lines.push("    For i = 1 To columnsUsed");
  lines.push("        wantedHeader = CellText(headers(1, i))");
  lines.push("        columnMap(i) = ColumnIndexByHeader(targetTable.HeaderRowRange, wantedHeader)");
  lines.push("        If columnMap(i) = 0 Then");
  lines.push(
    '            Err.Raise ERR_BASE + 12, MACRO_LABEL, "Destination table """ & targetTable.Name & """ has no column headed """ & wantedHeader & """, so this macro will not write to it -- putting the value under a different heading would corrupt the report silently. The table headings are: " & TableHeaderList(targetTable) & ". Either rename a table column to """ & wantedHeader & """, or point this macro at a different destination."'
  );
  lines.push("        End If");
  lines.push("    Next i");
  lines.push("End Sub");
  lines.push("");

  lines.push("' Readable list of a table's headings, used only in error messages.");
  lines.push("Private Function TableHeaderList(ByVal targetTable As ListObject) As String");
  lines.push("    Dim i As Long");
  lines.push("    Dim parts As String");
  lines.push("    If targetTable.HeaderRowRange Is Nothing Then Exit Function");
  lines.push("    For i = 1 To targetTable.HeaderRowRange.Cells.Count");
  lines.push("        If Len(parts) > 0 Then parts = parts & \", \"");
  lines.push("        parts = parts & \"\"\"\" & CellText(targetTable.HeaderRowRange.Cells(1, i).Value) & \"\"\"\"");
  lines.push("    Next i");
  lines.push("    TableHeaderList = parts");
  lines.push("End Function");
  lines.push("");

  lines.push("' Cell value as trimmed text. Error and empty cells become an empty");
  lines.push("' string rather than blowing up the caller.");
  lines.push("Private Function CellText(ByVal cellValue As Variant) As String");
  lines.push("    If IsError(cellValue) Then Exit Function");
  lines.push("    If IsEmpty(cellValue) Then Exit Function");
  lines.push("    If IsNull(cellValue) Then Exit Function");
  lines.push("    If IsObject(cellValue) Then Exit Function");
  lines.push("    CellText = Trim$(CStr(cellValue))");
  lines.push("End Function");
  lines.push("");

  lines.push("' Reads a range into a 1-based two dimensional array, including the");
  lines.push("' single-cell case, where Range.Value would give a bare scalar.");
  lines.push("Private Function ToArray2D(ByVal target As Range) As Variant");
  lines.push("    Dim result() As Variant");
  lines.push("    If target.Cells.Count = 1 Then");
  lines.push("        ReDim result(1 To 1, 1 To 1)");
  lines.push("        result(1, 1) = target.Value");
  lines.push("        ToArray2D = result");
  lines.push("    Else");
  lines.push("        ToArray2D = target.Value");
  lines.push("    End If");
  lines.push("End Function");
  lines.push("");

  if (plan.kind === "deduplicate" || plan.kind === "aggregate") {
    lines.push("' True when the collection already holds this key.");
    lines.push("Private Function CollectionHasKey(ByVal target As Collection, ByVal keyName As String) As Boolean");
    lines.push("    Dim probe As Variant");
    lines.push("    On Error Resume Next");
    lines.push("    probe = target.Item(keyName)");
    lines.push("    CollectionHasKey = (Err.Number = 0)");
    lines.push("    Err.Clear");
    lines.push("    On Error GoTo 0");
    lines.push("End Function");
    lines.push("");
  }

  if (needsSort(plan)) {
    lines.push("' Stable insertion sort of the first rowsUsed rows, ascending, by one");
    lines.push("' column, compared case-insensitively as text.");
    lines.push(
      "Private Sub SortRowsByColumn(ByRef target() As Variant, ByVal rowsUsed As Long, ByVal sortColumn As Long)"
    );
    lines.push("    Dim i As Long");
    lines.push("    Dim j As Long");
    lines.push("    Dim col As Long");
    lines.push("    Dim columnsUsed As Long");
    lines.push("    Dim liftedRow() As Variant");
    lines.push("    Dim liftedKey As String");
    lines.push("");
    lines.push("    If rowsUsed < 2 Then Exit Sub");
    lines.push("    columnsUsed = UBound(target, 2)");
    lines.push("    ReDim liftedRow(1 To columnsUsed)");
    lines.push("");
    lines.push("    For i = 2 To rowsUsed");
    lines.push("        For col = 1 To columnsUsed");
    lines.push("            liftedRow(col) = target(i, col)");
    lines.push("        Next col");
    lines.push("        liftedKey = CellText(liftedRow(sortColumn))");
    lines.push("        j = i - 1");
    lines.push("        Do While j >= 1");
    lines.push("            If StrComp(CellText(target(j, sortColumn)), liftedKey, vbTextCompare) <= 0 Then Exit Do");
    lines.push("            For col = 1 To columnsUsed");
    lines.push("                target(j + 1, col) = target(j, col)");
    lines.push("            Next col");
    lines.push("            j = j - 1");
    lines.push("        Loop");
    lines.push("        For col = 1 To columnsUsed");
    lines.push("            target(j + 1, col) = liftedRow(col)");
    lines.push("        Next col");
    lines.push("    Next i");
    lines.push("End Sub");
    lines.push("");
  }

  if (needsNumberParser(plan)) {
    lines.push("' Parses a cell as a number, tolerating the formatting a spreadsheet");
    lines.push("' shows: currency symbols, thousands separators, padding, and");
    lines.push("' parenthesised negatives such as (1,250.00). Returns False for text,");
    lines.push("' blanks, booleans, dates and error values -- the caller must then SKIP");
    lines.push("' the value, never treat it as zero. Val() is used instead of CDbl so");
    lines.push("' the parse cannot fail on unexpected text or a locale difference.");
    lines.push("Private Function TryParseNumber(ByVal cellValue As Variant, ByRef outValue As Double) As Boolean");
    lines.push("    Dim rawText As String");
    lines.push("    Dim cleaned As String");
    lines.push("    Dim ch As String");
    lines.push("    Dim i As Long");
    lines.push("    Dim negative As Boolean");
    lines.push("    Dim decimalSeen As Boolean");
    lines.push("");
    lines.push("    outValue = 0");
    lines.push("");
    lines.push("    Select Case VarType(cellValue)");
    lines.push("        Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal, vbByte");
    lines.push("            outValue = CDbl(cellValue)");
    lines.push("            TryParseNumber = True");
    lines.push("            Exit Function");
    lines.push("        Case vbString");
    lines.push("            ' Handled by the text parsing below.");
    lines.push("        Case Else");
    lines.push("            ' Empty, Null, Boolean, Date, Error, Object: not a number.");
    lines.push("            Exit Function");
    lines.push("    End Select");
    lines.push("");
    lines.push("    rawText = Trim$(CStr(cellValue))");
    lines.push("    If Len(rawText) = 0 Then Exit Function");
    lines.push("");
    lines.push('    If Left$(rawText, 1) = "(" And Right$(rawText, 1) = ")" Then');
    lines.push("        negative = True");
    lines.push("        rawText = Trim$(Mid$(rawText, 2, Len(rawText) - 2))");
    lines.push("    End If");
    lines.push("");
    lines.push("    For i = 1 To Len(rawText)");
    lines.push("        ch = Mid$(rawText, i, 1)");
    lines.push('        If ch = "$" Or ch = "," Or ch = " " Or ch = vbTab _');
    lines.push("           Or ch = ChrW$(163) Or ch = ChrW$(8364) Then");
    lines.push("            ' Currency symbol, thousands separator or padding: ignored.");
    lines.push("        Else");
    lines.push("            cleaned = cleaned & ch");
    lines.push("        End If");
    lines.push("    Next i");
    lines.push("");
    lines.push('    If Left$(cleaned, 1) = "-" Then');
    lines.push("        negative = Not negative");
    lines.push("        cleaned = Mid$(cleaned, 2)");
    lines.push('    ElseIf Left$(cleaned, 1) = "+" Then');
    lines.push("        cleaned = Mid$(cleaned, 2)");
    lines.push("    End If");
    lines.push("");
    lines.push("    If Len(cleaned) = 0 Then Exit Function");
    lines.push("    ' Digits with at most one decimal point, ending in a digit.");
    lines.push("    For i = 1 To Len(cleaned)");
    lines.push("        ch = Mid$(cleaned, i, 1)");
    lines.push('        If ch = "." Then');
    lines.push("            If decimalSeen Then Exit Function");
    lines.push("            decimalSeen = True");
    lines.push('        ElseIf ch < "0" Or ch > "9" Then');
    lines.push("            Exit Function");
    lines.push("        End If");
    lines.push("    Next i");
    lines.push('    If Right$(cleaned, 1) = "." Then Exit Function');
    lines.push("");
    lines.push("    outValue = Val(cleaned)");
    lines.push("    If negative Then outValue = -outValue");
    lines.push("    TryParseNumber = True");
    lines.push("End Function");
    lines.push("");
  }

  return lines;
}

function buildVba(plan: Plan, rules: UnimplementedRule[], steps: string[]): string {
  const lines: string[] = [];
  lines.push("Option Explicit");
  lines.push("");
  lines.push(...buildHeaderBlock(plan, rules, steps));
  lines.push("");
  lines.push(...buildConfigBlock(plan));
  lines.push("");
  if (plan.trigger === "button") {
    lines.push("' Assign this Sub to your button (right-click the button, Assign Macro).");
  }
  lines.push(`Public Sub ${plan.macroName}()`);
  lines.push(...buildDimBlock(plan));
  lines.push("");
  lines.push("    On Error GoTo CleanFail");
  lines.push("");
  lines.push("    ' Save the Application state BEFORE changing it, so the cleanup path");
  lines.push("    ' below can always put Excel back the way it was.");
  lines.push("    prevScreenUpdating = Application.ScreenUpdating");
  lines.push("    prevEnableEvents = Application.EnableEvents");
  lines.push("    prevDisplayAlerts = Application.DisplayAlerts");
  lines.push("    prevCalculation = Application.Calculation");
  lines.push("    stateSaved = True");
  lines.push("");
  lines.push("    Application.ScreenUpdating = False");
  lines.push("    Application.EnableEvents = False");
  lines.push("    Application.Calculation = xlCalculationManual");
  lines.push("    ' DisplayAlerts is captured and restored, but deliberately left ON:");
  lines.push("    ' this macro never saves or deletes anything, so suppressing Excel's");
  lines.push("    ' own warnings would only hide problems from you.");
  lines.push("");
  lines.push(...buildReadSourceBlock(plan));
  lines.push("");
  lines.push(...buildTransformBlock(plan));
  lines.push("");
  lines.push(...buildDestinationBlock(plan));
  lines.push("");
  lines.push(...buildReportBlock(plan));
  lines.push("");
  lines.push("    GoTo Cleanup");
  lines.push("");
  lines.push("CleanFail:");
  lines.push("    failed = True");
  lines.push(
    '    failMessage = MACRO_LABEL & " stopped and changed nothing further." & vbCrLf & vbCrLf & _'
  );
  lines.push('        Err.Description & vbCrLf & "(error " & Err.Number & ")"');
  lines.push("");
  lines.push("Cleanup:");
  lines.push("    ' Always runs, on success and on failure, so an early exit can never");
  lines.push("    ' leave Excel with events off or calculation on manual. Errors are");
  lines.push("    ' ignored from here on so a failure while restoring cannot bounce");
  lines.push("    ' back into the handler and loop.");
  lines.push("    On Error Resume Next");
  lines.push("    If stateSaved Then");
  lines.push("        Application.Calculation = prevCalculation");
  lines.push("        Application.DisplayAlerts = prevDisplayAlerts");
  lines.push("        Application.EnableEvents = prevEnableEvents");
  lines.push("        Application.ScreenUpdating = prevScreenUpdating");
  lines.push("    End If");
  lines.push("");
  lines.push("    Set headerRow = Nothing");
  lines.push("    Set dataBody = Nothing");
  lines.push("    Set sourceRange = Nothing");
  lines.push("    Set sourceSheet = Nothing");
  lines.push("    Set sourceBook = Nothing");
  lines.push("    Set destAnchor = Nothing");
  lines.push("    Set destTable = Nothing");
  lines.push("    Set destSheet = Nothing");
  lines.push("    Set destBook = Nothing");
  lines.push("");
  lines.push("    If failed Then");
  lines.push('        MsgBox failMessage, vbExclamation, "Excel Macro Builder"');
  lines.push("    Else");
  lines.push('        MsgBox reportText, vbInformation, "Excel Macro Builder"');
  lines.push("    End If");
  lines.push("End Sub");
  lines.push("");
  lines.push(...buildHelpers(plan));

  return lines.join("\n").replace(/\n{3,}/g, "\n\n").trimEnd() + "\n";
}

// ---------------------------------------------------------------------------
// Narrative
// ---------------------------------------------------------------------------

function buildSteps(plan: Plan): string[] {
  const steps: string[] = [];
  const sourceLabel = `"${plan.sourceRangeOrTable}" on "${plan.sourceWorksheet}" in ${plan.sourceWorkbook}`;
  steps.push(`Read the rows of ${sourceLabel}, treating its first row as the header row.`);

  switch (plan.kind) {
    case "copy":
      steps.push("Copy every row through unchanged.");
      break;
    case "filter": {
      const label = FILTER_LABELS[plan.filterOperator];
      const needsValue = plan.filterOperator !== "is-blank" && plan.filterOperator !== "is-not-blank";
      steps.push(
        needsValue
          ? `Keep only rows where "${plan.keyColumn}" ${label} "${plan.filterValue}".`
          : `Keep only rows where "${plan.keyColumn}" ${label}.`
      );
      break;
    }
    case "deduplicate":
      steps.push(`Keep the FIRST row for each distinct "${plan.keyColumn}" value.`);
      steps.push(`Sort the result by "${plan.keyColumn}" ascending.`);
      break;
    case "aggregate":
      steps.push(`Group the rows by "${plan.keyColumn}".`);
      steps.push(
        plan.aggregate === "count"
          ? "Count the rows in each group."
          : `Take the ${AGGREGATE_LABELS[plan.aggregate].toLowerCase()} of "${plan.valueColumn}" within each group, skipping blank and non-numeric values instead of counting them as zero.`
      );
      steps.push(`Sort the result by "${plan.keyColumn}" ascending.`);
      break;
  }

  const destBook = plan.sameWorkbook ? "the same workbook" : plan.destinationWorkbook;
  steps.push(
    plan.overwrite
      ? `Clear only the destination block "${plan.destinationRangeOrTable}" on "${plan.destinationWorksheet}" in ${destBook}, then write the result there.`
      : `Append the result underneath the existing rows of "${plan.destinationRangeOrTable}" on "${plan.destinationWorksheet}" in ${destBook}.`
  );
  steps.push("Restore Excel's Application settings and report what was written in a message box.");
  return steps;
}

function buildAssumptions(plan: Plan): string[] {
  const assumptions: string[] = [
    "The first row of the source range or table is the header row; all columns are found by header name, never by position.",
    "The source workbook is already open in Excel. The macro reports a clear error rather than opening it for you.",
    "Key comparison, grouping and sorting are case-insensitive and ignore surrounding spaces, matching the on-screen preview.",
    "Sorting uses VBA's case-insensitive text comparison, which can order accented or non-Latin text slightly differently from the browser preview.",
  ];

  if (plan.kind === "aggregate" && plan.aggregate !== "count") {
    assumptions.push(
      "Blank and non-numeric values in the aggregated column are skipped, never treated as zero, and the count of skipped values is reported at the end."
    );
    assumptions.push(
      'A group with no numeric values at all is written as an em dash ("—"), not as 0, because there is no honest total to report.'
    );
  }

  if (plan.sameWorkbook) {
    assumptions.push("The source and destination are the same workbook, so no second workbook is opened.");
  } else {
    assumptions.push(
      "If the destination workbook is already open it is reused; otherwise it is opened from the configured path."
    );
    assumptions.push(
      "The destination workbook is deliberately NOT saved and NOT closed by the macro. It is left open so you can inspect the result and decide whether to keep it."
    );
  }

  assumptions.push(
    plan.overwrite
      ? "Overwrite clears only the destination table's data body, or the rectangle from the anchor cell down across the columns written. It never clears the whole sheet and never deletes rows."
      : "Append adds rows underneath what is already there and clears nothing."
  );

  assumptions.push(
    "The destination table or range is assumed to have at least as many columns as the macro writes; if it does not, the macro stops with a clear error instead of writing part of the result."
  );
  assumptions.push(
    `When the destination is an Excel table, each value is written to the table column whose heading matches the result heading (${outputHeaderNames(
      plan
    )
      .map((h) => `"${h}"`)
      .join(
        ", "
      )}), not to whichever column happens to sit in that position. The macro does not rename your table's headings. If a heading is missing it stops before clearing anything and tells you which headings the table actually has, because filing a value under the wrong heading would corrupt the report silently.`
  );
  assumptions.push(
    "The macro finishes with a message box. If your constraints say it must not prompt, remove that MsgBox yourself."
  );

  return assumptions;
}

function buildTestPlan(plan: Plan): string[] {
  return [
    `Make a COPY of ${plan.sourceWorkbook}${plan.sameWorkbook ? "" : ` and ${plan.destinationWorkbook}`} and work only on the copy.`,
    "Read the whole macro before running it, especially the CONFIGURATION block and the not-implemented comment block at the top.",
    `Run it once with the destination "${plan.destinationRangeOrTable}" empty and confirm the header row and results land where you expect.`,
    plan.overwrite
      ? "Run it a second time and confirm the previous result is replaced, and that nothing outside the destination block changed."
      : "Run it a second time and confirm the new rows are appended below the previous ones, with nothing overwritten.",
    "Rename a source column header temporarily and confirm the macro stops with an error naming the missing header, rather than writing wrong data.",
    "Point the source at an empty range and confirm the macro reports that there are no data rows.",
    plan.kind === "aggregate" && plan.aggregate !== "count"
      ? "Put text (e.g. \"n/a\") and a blank in the value column, then confirm the final message box reports them as skipped and the totals exclude them."
      : "Compare the written result against the before/after preview in the app for the same sample rows.",
    "Check Excel afterwards: calculation should be back to Automatic and events should be back on.",
  ];
}

function buildSafetyCautions(plan: Plan): string[] {
  const cautions: string[] = [
    "This code has not been run, compiled, or verified by anything. Deterministic does not mean correct, and it certainly does not mean safe.",
    "The template implements ONLY the structured operation you picked. Every free-text rule you typed is listed as not implemented and is genuinely absent from the code.",
    "The app has never checked that these workbooks, sheets, tables or headers exist. Those names are text you typed.",
  ];
  if (plan.overwrite) {
    cautions.push(
      "This macro writes in overwrite mode: existing values in the destination block are cleared. Confirm the destination is what you think it is before running it on real data."
    );
  }
  if (!plan.sameWorkbook) {
    cautions.push(
      "The macro may open the destination workbook. It never saves or closes it, so nothing is written to disk unless you save it yourself."
    );
  }
  return cautions;
}

function buildPlatformLimitations(form: MacroFormData, plan: Plan): string[] {
  const limitations: string[] = [
    "The generated code avoids Windows-only features: no Scripting.Dictionary, no ActiveX, no Windows file dialogs. It uses a VBA Collection so it works on Mac Excel too.",
  ];
  if (!plan.sameWorkbook) {
    limitations.push(
      "Path separators differ between Windows (\\) and Mac (/). Set DEST_WORKBOOK_PATH to a full path in the form your Excel uses, or open the destination workbook first."
    );
    limitations.push(
      "On Mac Excel, opening a file by path can require granting file access the first time. Opening the destination workbook manually avoids that."
    );
  }
  if (form.task.platform === "cross-platform") {
    limitations.push(
      "You selected cross-platform. Test on both Windows and Mac Excel before relying on it in either."
    );
  }
  limitations.push(
    "The macro must live in a standard module of a macro-enabled workbook (.xlsm). It cannot be saved in a plain .xlsx file."
  );
  return limitations;
}

function buildInstallInstructions(plan: Plan): string {
  const steps: string[] = [
    "1. Open Excel and open the VBA editor (Alt+F11 on Windows; Tools > Macro > Visual Basic Editor on Mac).",
    "2. Insert > Module to create a standard module, and paste this code into it.",
    "3. Check the CONFIGURATION block at the top: workbook names/paths, sheet names, the table or range, and the header names.",
    "4. Read the NOT IMPLEMENTED comment block and add anything it lists that you actually need.",
    "5. Save the workbook as a macro-enabled file (.xlsm).",
  ];
  steps.push(
    plan.sameWorkbook
      ? `6. Open ${plan.sourceWorkbook}.`
      : `6. Open ${plan.sourceWorkbook}. The destination workbook can be open or closed; the macro reuses it if it is open.`
  );
  steps.push(
    plan.trigger === "button"
      ? `7. Add a button (Developer > Insert > Button), right-click it, choose Assign Macro, and pick ${plan.macroName}.`
      : `7. Run ${plan.macroName} from the VBA editor (F5) or from Developer > Macros.`
  );
  steps.push("8. Do all of the above on a COPY of your workbook first.");
  return steps.join("\n");
}

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------

/**
 * Builds real VBA from the structured step-2 picker, with no AI and no network.
 *
 * Refuses (status "unsupported", empty `vbaCode`) rather than guessing whenever
 * the request falls outside the four structured operations, the two supported
 * triggers, or the fields the template actually needs. Never throws.
 */
export function generateVbaFromTemplate(form: MacroFormData): TemplateVbaResult {
  if (!form || !form.mapping || !form.task || !form.run) {
    return unsupported("The form data is incomplete, so no macro could be built from it.");
  }

  const preview = form.preview;
  if (!preview || preview.kind === "not-configured") {
    return unsupported(
      'No preview operation is configured. Go back to step 2 and pick one under "Preview operation" (copy rows, filter rows, remove duplicates, or group and total). The deterministic builder works from that picker, so it has nothing to build from until you choose one.'
    );
  }

  const trigger = form.run.trigger;
  if (trigger === "workbook-open" || trigger === "sheet-change") {
    const where =
      trigger === "workbook-open"
        ? "a Workbook_Open handler in the ThisWorkbook code module"
        : "a Worksheet_Change handler in that sheet's own code module";
    return unsupported(
      `The "${trigger === "workbook-open" ? "On workbook open" : "On sheet change"}" trigger needs ${where}, not the standard module this builder emits. Event handlers also fire on their own, which needs care this template does not attempt. Switch the trigger to "Run manually" or "From a button" to use the deterministic builder, or use the AI path (or hand-author the handler) if you really need the event trigger.`
    );
  }

  const nameCheck = validateVbaMacroName(form.task.macroName ?? "");
  if (!nameCheck.valid) {
    return unsupported(`The macro name cannot be used as a VBA Sub name: ${nameCheck.errors.join(" ")}`);
  }

  const required: [string, string][] = [
    ["Source workbook", form.mapping.sourceWorkbook],
    ["Source worksheet", form.mapping.sourceWorksheet],
    ["Source range or table", form.mapping.sourceRangeOrTable],
    ["Relevant column headers", form.mapping.sourceColumnHeaders],
    ["Destination worksheet", form.mapping.destinationWorksheet],
    ["Destination table or starting cell", form.mapping.destinationRangeOrTable],
  ];
  const missing = required.filter(([, value]) => (value ?? "").trim().length === 0).map(([label]) => label);
  if (missing.length > 0) {
    return unsupported(
      `These step 2 fields are needed before real code can be built, and the template will not invent them: ${missing.join(", ")}.`
    );
  }

  const sameWorkbook = form.mapping.sameWorkbook;
  let destinationWorkbook = "";
  if (!sameWorkbook) {
    const dest = filledMaybe(form.mapping.destinationWorkbook);
    if (dest === null) {
      return unsupported(
        "Source and destination are different workbooks, so a destination workbook file name or path is required. Fill it in on step 2, or tick \"Source and destination are the same workbook\"."
      );
    }
    destinationWorkbook = dest;
  }

  const headers = parseSampleData(form.mapping.sourceColumnHeaders, "").headers;
  if (headers.length === 0) {
    return unsupported(
      'No column headers could be read from "Relevant column headers" on step 2. Enter them as a comma-separated list, e.g. "Account Name, Account Number, Amount".'
    );
  }

  function findHeader(name: string): string | null {
    const target = (name ?? "").trim().toLowerCase();
    if (target.length === 0) return null;
    return headers.find((h) => h.toLowerCase() === target) ?? null;
  }

  let keyColumn = "";
  if (preview.kind !== "copy") {
    const label = preview.kind === "filter" ? "Column to test" : "Group/key column";
    const matched = findHeader(preview.keyColumn);
    if (matched === null) {
      const chosen = (preview.keyColumn ?? "").trim();
      return unsupported(
        chosen.length === 0
          ? `The preview operation needs a ${label.toLowerCase()}, and none is set. Pick one on step 2.`
          : `${label} "${chosen}" is not one of the source column headers (${headers.join(", ")}). Fix it on step 2 — the template will not guess which column you meant.`
      );
    }
    keyColumn = matched;
  }

  let valueColumn = "";
  if (preview.kind === "aggregate" && preview.aggregate !== "count") {
    const matched = findHeader(preview.valueColumn);
    if (matched === null) {
      const chosen = (preview.valueColumn ?? "").trim();
      return unsupported(
        chosen.length === 0
          ? "This aggregate needs a numeric value column, and none is set. Pick one on step 2."
          : `Value column "${chosen}" is not one of the source column headers (${headers.join(", ")}). Fix it on step 2 — the template will not guess which column you meant.`
      );
    }
    valueColumn = matched;
  }

  const plan: Plan = {
    macroName: form.task.macroName.trim(),
    sourceWorkbook: form.mapping.sourceWorkbook.trim(),
    sourceWorksheet: form.mapping.sourceWorksheet.trim(),
    sourceRangeOrTable: form.mapping.sourceRangeOrTable.trim(),
    sameWorkbook,
    destinationWorkbook,
    destinationWorksheet: form.mapping.destinationWorksheet.trim(),
    destinationRangeOrTable: form.mapping.destinationRangeOrTable.trim(),
    // "not-applicable" writes without clearing, which is what append does.
    overwrite: form.mapping.appendOrOverwrite === "overwrite",
    kind: preview.kind,
    keyColumn,
    valueColumn,
    aggregate: preview.aggregate,
    filterOperator: preview.filterOperator,
    filterValue: (preview.filterValue ?? "").trim(),
    headers,
    trigger,
  };

  const unimplementedRules = collectUnimplementedRules(form);
  const steps = buildSteps(plan);
  const vbaCode = buildVba(plan, unimplementedRules, steps);

  const operationWord =
    plan.kind === "aggregate"
      ? `${AGGREGATE_LABELS[plan.aggregate].toLowerCase()} by "${plan.keyColumn}"`
      : plan.kind === "filter"
        ? `filter on "${plan.keyColumn}"`
        : plan.kind === "deduplicate"
          ? `deduplicate by "${plan.keyColumn}"`
          : "straight copy";

  const summary =
    `Template-generated VBA (no AI): ${operationWord} from "${plan.sourceRangeOrTable}" on ` +
    `"${plan.sourceWorksheet}", written to "${plan.destinationRangeOrTable}" on "${plan.destinationWorksheet}"` +
    `${plan.sameWorkbook ? " in the same workbook" : ` in ${plan.destinationWorkbook}`} in ` +
    `${plan.overwrite ? "overwrite" : "append"} mode. Built from the same structured operation as the on-screen ` +
    `preview, so the two cannot disagree. It implements nothing else.`;

  return {
    status: "ok",
    vbaCode,
    summary,
    steps,
    unimplementedRules,
    assumptions: buildAssumptions(plan),
    installInstructions: buildInstallInstructions(plan),
    testPlan: buildTestPlan(plan),
    safetyCautions: buildSafetyCautions(plan),
    platformLimitations: buildPlatformLimitations(form, plan),
  };
}
