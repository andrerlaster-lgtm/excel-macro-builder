import { AggregateFn, FilterOperator, PreviewConfig } from "./types";

// ---------------------------------------------------------------------------
// Deterministic in-browser preview.
//
// This module simulates the operation the user DESCRIBED via the structured
// picker on step 2. It never sees, parses, or executes the generated VBA, so
// it cannot prove the real macro behaves this way -- it only shows what the
// described operation would do to the sample rows. Every caller must present
// it with that caveat attached.
//
// Pure functions only: no DOM, no network, no clock.
// ---------------------------------------------------------------------------

/** Hard caps so a huge paste can never lock up the browser. */
export const MAX_PREVIEW_ROWS = 500;
export const MAX_PREVIEW_COLUMNS = 50;

export interface PreviewTable {
  headers: string[];
  rows: string[][];
}

export interface PreviewOutcome {
  status: "ok" | "not-configured" | "error";
  /** Why it couldn't simulate (missing column, no rows, etc.). */
  message?: string;
  before: PreviewTable;
  after: PreviewTable;
  /** Human-readable narrative of each stage applied, in order. */
  steps: string[];
  /** Caveats, e.g. non-numeric values skipped. */
  notes: string[];
}

const EMPTY_TABLE: PreviewTable = { headers: [], rows: [] };

function emptyOutcome(status: PreviewOutcome["status"], message: string, before: PreviewTable = EMPTY_TABLE): PreviewOutcome {
  return { status, message, before, after: EMPTY_TABLE, steps: [], notes: [] };
}

/**
 * Splits the free-text step 2 fields into a grid.
 *
 * Headers come from the comma-separated "Relevant column headers" field;
 * rows come from "Example rows", one row per line, comma separated. Rows are
 * padded or trimmed to the header count so the grid is always rectangular.
 */
export function parseSampleData(
  columnHeaders: string,
  exampleRows: string
): { headers: string[]; rows: string[][] } {
  const headers = (columnHeaders ?? "")
    .split(",")
    .map((h) => h.trim())
    .filter((h) => h.length > 0)
    .slice(0, MAX_PREVIEW_COLUMNS);

  if (headers.length === 0) return { headers: [], rows: [] };

  const rows = (exampleRows ?? "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0)
    .slice(0, MAX_PREVIEW_ROWS)
    .map((line) => {
      const cells = line.split(",").map((c) => c.trim());
      // Drop a single trailing empty cell caused by a trailing comma.
      if (cells.length > headers.length && cells[cells.length - 1] === "") cells.pop();

      if (cells.length > headers.length) {
        // Never silently delete the tail of an over-long row: a value like
        // "$1,200.00" splits on its own thousands separator, and dropping the
        // remainder would turn 1200 into 1 and produce a confidently wrong
        // total. Fold the overflow back into the last cell instead, where it
        // stays visible and editable in the preview grid.
        const kept = cells.slice(0, headers.length - 1);
        kept.push(cells.slice(headers.length - 1).join(","));
        return kept;
      }

      const padded = [...cells];
      while (padded.length < headers.length) padded.push("");
      return padded;
    });

  return { headers, rows };
}

/**
 * Parses a cell as a number, forgiving common spreadsheet formatting
 * ($, thousands separators, surrounding whitespace, parenthesised negatives).
 * Returns null when the cell is not a number -- callers must skip and report
 * it rather than coercing junk to zero.
 */
export function parseNumericCell(raw: string): number | null {
  const trimmed = (raw ?? "").trim();
  if (trimmed.length === 0) return null;

  let working = trimmed;
  let negative = false;

  if (/^\(.*\)$/.test(working)) {
    negative = true;
    working = working.slice(1, -1).trim();
  }

  working = working.replace(/[$£€]/g, "").replace(/,/g, "").replace(/\s/g, "");

  if (working.startsWith("-")) {
    negative = !negative;
    working = working.slice(1);
  } else if (working.startsWith("+")) {
    working = working.slice(1);
  }

  if (!/^\d*\.?\d+$/.test(working)) return null;

  const value = Number(working);
  if (!Number.isFinite(value)) return null;
  return negative ? -value : value;
}

/** Trims trailing zeros so 120 stays "120" and 120.5 stays "120.5". */
function formatNumber(value: number): string {
  if (Number.isInteger(value)) return String(value);
  return String(Number(value.toFixed(4)));
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

export function aggregateLabel(fn: AggregateFn): string {
  return AGGREGATE_LABELS[fn];
}

export function filterOperatorLabel(op: FilterOperator): string {
  return FILTER_LABELS[op];
}

function matchesFilter(cell: string, operator: FilterOperator, filterValue: string): boolean {
  const value = (cell ?? "").trim();
  const target = (filterValue ?? "").trim();

  switch (operator) {
    case "is-blank":
      return value.length === 0;
    case "is-not-blank":
      return value.length > 0;
    case "equals":
      return value.toLowerCase() === target.toLowerCase();
    case "not-equals":
      return value.toLowerCase() !== target.toLowerCase();
    case "contains":
      return value.toLowerCase().includes(target.toLowerCase());
    case "greater-than":
    case "less-than": {
      const left = parseNumericCell(value);
      const right = parseNumericCell(target);
      if (left !== null && right !== null) {
        return operator === "greater-than" ? left > right : left < right;
      }
      // Fall back to case-insensitive text ordering when either side is not
      // numeric, so date-like or text values still compare sensibly.
      const cmp = value.toLowerCase().localeCompare(target.toLowerCase());
      return operator === "greater-than" ? cmp > 0 : cmp < 0;
    }
    default:
      return true;
  }
}

/**
 * Simulates the configured operation against the sample grid.
 *
 * Stages are applied in a fixed order: filter, then deduplicate, then
 * aggregate. Aggregate and deduplicate output is sorted by key ascending so
 * the result is stable and comparable between edits.
 *
 * Never throws: unusable configurations come back as a status plus a message.
 */
export function simulatePreview(config: PreviewConfig): PreviewOutcome {
  if (!config || config.kind === "not-configured") {
    return emptyOutcome(
      "not-configured",
      "Pick a preview operation in step 2 to see a before/after example."
    );
  }

  const notes: string[] = [];
  const steps: string[] = [];

  let headers = (config.sampleHeaders ?? []).map((h) => (h ?? "").trim());
  if (headers.length > MAX_PREVIEW_COLUMNS) {
    headers = headers.slice(0, MAX_PREVIEW_COLUMNS);
    notes.push(`Only the first ${MAX_PREVIEW_COLUMNS} columns are simulated.`);
  }

  if (headers.length === 0) {
    return emptyOutcome(
      "not-configured",
      "Add column headers in step 2 (or in the sample grid below) to see a before/after example."
    );
  }

  let sourceRows = (config.sampleRows ?? []).map((row) => {
    const cells = (row ?? []).slice(0, headers.length).map((c) => c ?? "");
    while (cells.length < headers.length) cells.push("");
    return cells;
  });

  if (sourceRows.length > MAX_PREVIEW_ROWS) {
    sourceRows = sourceRows.slice(0, MAX_PREVIEW_ROWS);
    notes.push(`Only the first ${MAX_PREVIEW_ROWS} rows are simulated.`);
  }

  const before: PreviewTable = { headers: [...headers], rows: sourceRows.map((r) => [...r]) };

  if (sourceRows.length === 0) {
    return {
      ...emptyOutcome("not-configured", "Add at least one sample row to see a before/after example.", before),
      notes,
    };
  }

  function columnIndex(name: string): number {
    const target = (name ?? "").trim().toLowerCase();
    if (target.length === 0) return -1;
    return headers.findIndex((h) => h.toLowerCase() === target);
  }

  // --- copy -----------------------------------------------------------------
  if (config.kind === "copy") {
    steps.push(`Copy all ${sourceRows.length} row(s) through unchanged.`);
    return {
      status: "ok",
      before,
      after: { headers: [...headers], rows: sourceRows.map((r) => [...r]) },
      steps,
      notes,
    };
  }

  const keyIndex = columnIndex(config.keyColumn);
  if (keyIndex === -1) {
    const label = config.kind === "filter" ? "Column to test" : "Key column";
    return {
      ...emptyOutcome(
        "error",
        `${label} "${config.keyColumn.trim() || "(not set)"}" is not one of the sample columns (${headers.join(", ")}).`,
        before
      ),
      notes,
    };
  }

  // --- filter ---------------------------------------------------------------
  if (config.kind === "filter") {
    const opLabel = filterOperatorLabel(config.filterOperator);
    const needsValue = config.filterOperator !== "is-blank" && config.filterOperator !== "is-not-blank";
    const kept = sourceRows.filter((row) => matchesFilter(row[keyIndex], config.filterOperator, config.filterValue));
    steps.push(
      needsValue
        ? `Keep rows where "${headers[keyIndex]}" ${opLabel} "${config.filterValue.trim()}" — ${kept.length} of ${sourceRows.length} row(s) kept.`
        : `Keep rows where "${headers[keyIndex]}" ${opLabel} — ${kept.length} of ${sourceRows.length} row(s) kept.`
    );
    if (kept.length === 0) {
      notes.push("No sample row matched this filter, so the destination would be empty.");
    }
    return { status: "ok", before, after: { headers: [...headers], rows: kept }, steps, notes };
  }

  // --- deduplicate ----------------------------------------------------------
  if (config.kind === "deduplicate") {
    const seen = new Set<string>();
    const kept: string[][] = [];
    for (const row of sourceRows) {
      const key = row[keyIndex].trim().toLowerCase();
      if (seen.has(key)) continue;
      seen.add(key);
      kept.push([...row]);
    }
    kept.sort((a, b) => a[keyIndex].trim().localeCompare(b[keyIndex].trim()));
    steps.push(
      `Keep the FIRST row for each distinct "${headers[keyIndex]}" value — ${kept.length} of ${sourceRows.length} row(s) kept.`
    );
    steps.push(`Sort the result by "${headers[keyIndex]}" ascending.`);
    steps.push(
      'This simulation always keeps the first occurrence. Your free-text "Duplicate handling" rule may say something different (e.g. keep the most recent), and the generated macro follows that rule, not this one.'
    );
    return { status: "ok", before, after: { headers: [...headers], rows: kept }, steps, notes };
  }

  // --- aggregate ------------------------------------------------------------
  const fn = config.aggregate;
  const needsValueColumn = fn !== "count";
  let valueIndex = -1;

  if (needsValueColumn) {
    valueIndex = columnIndex(config.valueColumn);
    if (valueIndex === -1) {
      return {
        ...emptyOutcome(
          "error",
          `Value column "${config.valueColumn.trim() || "(not set)"}" is not one of the sample columns (${headers.join(", ")}).`,
          before
        ),
        notes,
      };
    }
  }

  const groups = new Map<string, { label: string; numbers: number[]; count: number }>();
  const skipped: string[] = [];

  for (const row of sourceRows) {
    const rawKey = row[keyIndex].trim();
    const groupKey = rawKey.toLowerCase();
    let group = groups.get(groupKey);
    if (!group) {
      group = { label: rawKey, numbers: [], count: 0 };
      groups.set(groupKey, group);
    }
    group.count += 1;

    if (needsValueColumn) {
      const parsed = parseNumericCell(row[valueIndex]);
      if (parsed === null) {
        if (row[valueIndex].trim().length > 0) skipped.push(row[valueIndex].trim());
        else skipped.push("(blank)");
        continue;
      }
      group.numbers.push(parsed);
    }
  }

  const valueHeader = needsValueColumn
    ? `${AGGREGATE_LABELS[fn]} of ${headers[valueIndex]}`
    : "Count of rows";

  const outRows: string[][] = [];
  for (const group of groups.values()) {
    let cell: string;
    if (!needsValueColumn) {
      cell = String(group.count);
    } else if (group.numbers.length === 0) {
      cell = "—";
    } else {
      const nums = group.numbers;
      const value =
        fn === "sum"
          ? nums.reduce((a, b) => a + b, 0)
          : fn === "average"
            ? nums.reduce((a, b) => a + b, 0) / nums.length
            : fn === "min"
              ? Math.min(...nums)
              : Math.max(...nums);
      cell = formatNumber(value);
    }
    outRows.push([group.label, cell]);
  }
  outRows.sort((a, b) => a[0].localeCompare(b[0]));

  steps.push(`Group the ${sourceRows.length} row(s) by "${headers[keyIndex]}" — ${outRows.length} distinct value(s).`);
  steps.push(
    needsValueColumn
      ? `Take the ${AGGREGATE_LABELS[fn].toLowerCase()} of "${headers[valueIndex]}" within each group.`
      : "Count the rows in each group."
  );
  steps.push(`Sort the result by "${headers[keyIndex]}" ascending.`);

  if (skipped.length > 0) {
    const unique = Array.from(new Set(skipped)).slice(0, 5);
    notes.push(
      `${skipped.length} value(s) in "${headers[valueIndex]}" were not numeric and were skipped, not counted as zero: ${unique
        .map((v) => `"${v}"`)
        .join(", ")}${unique.length < new Set(skipped).size ? ", …" : ""}`
    );
  }
  if (outRows.some((r) => r[1] === "—")) {
    notes.push('Groups shown as "—" had no numeric values at all, so no total could be computed.');
  }

  return {
    status: "ok",
    before,
    after: { headers: [headers[keyIndex], valueHeader], rows: outRows },
    steps,
    notes,
  };
}
