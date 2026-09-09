import { describe, expect, it } from "vitest";
import { emptyPreviewConfig, PreviewConfig } from "./types";
import {
  MAX_PREVIEW_COLUMNS,
  MAX_PREVIEW_ROWS,
  parseNumericCell,
  parseSampleData,
  simulatePreview,
} from "./previewSimulator";

function config(overrides: Partial<PreviewConfig>): PreviewConfig {
  return { ...emptyPreviewConfig(), ...overrides };
}

const HEADERS = ["Region", "Amount"];
const ROWS = [
  ["East", "100"],
  ["West", "50"],
  ["East", "20"],
  ["West", "70"],
];

describe("parseSampleData", () => {
  it("parses clean comma-separated headers and rows", () => {
    const { headers, rows } = parseSampleData("Date, Region, Amount", "2026-01-05, East, 120.00\n2026-01-06, West, 80");
    expect(headers).toEqual(["Date", "Region", "Amount"]);
    expect(rows).toEqual([
      ["2026-01-05", "East", "120.00"],
      ["2026-01-06", "West", "80"],
    ]);
  });

  it("ignores blank lines", () => {
    const { rows } = parseSampleData("A, B", "1, 2\n\n   \n3, 4\n");
    expect(rows).toEqual([
      ["1", "2"],
      ["3", "4"],
    ]);
  });

  it("pads short rows", () => {
    const { rows } = parseSampleData("A, B, C", "1");
    expect(rows).toEqual([["1", "", ""]]);
  });

  it("folds an over-long row's overflow into the last cell instead of deleting it", () => {
    const { rows } = parseSampleData("A, B, C", "1, 2, 3, 4, 5");
    expect(rows).toEqual([["1", "2", "3,4,5"]]);
  });

  it("keeps a currency value whose thousands separator is a comma", () => {
    // Regression: this used to silently truncate "$1,200.00" to "$1", which
    // made the preview report a total of 1 instead of 1200.
    const { rows } = parseSampleData("Date, Region, Amount", "2026-01-05, East, $1,200.00");
    expect(rows).toEqual([["2026-01-05", "East", "$1,200.00"]]);
  });

  it("tolerates a trailing comma on a row", () => {
    const { rows } = parseSampleData("A, B", "1, 2,");
    expect(rows).toEqual([["1", "2"]]);
  });

  it("returns empty structures for empty input", () => {
    expect(parseSampleData("", "")).toEqual({ headers: [], rows: [] });
    expect(parseSampleData("A, B", "")).toEqual({ headers: ["A", "B"], rows: [] });
  });
});

describe("parseNumericCell", () => {
  it("parses currency and thousands separators", () => {
    expect(parseNumericCell("$1,234.50")).toBe(1234.5);
    expect(parseNumericCell(" 1,000 ")).toBe(1000);
  });

  it("parses negatives and decimals", () => {
    expect(parseNumericCell("-12.5")).toBe(-12.5);
    expect(parseNumericCell("(300)")).toBe(-300);
    expect(parseNumericCell("0.25")).toBe(0.25);
  });

  it("returns null for non-numeric and blank cells", () => {
    expect(parseNumericCell("pending")).toBeNull();
    expect(parseNumericCell("")).toBeNull();
    expect(parseNumericCell("12abc")).toBeNull();
  });
});

describe("simulatePreview — statuses", () => {
  it("reports not-configured when no operation is picked", () => {
    const outcome = simulatePreview(emptyPreviewConfig());
    expect(outcome.status).toBe("not-configured");
    expect(outcome.message).toMatch(/step 2/i);
    expect(outcome.after.rows).toEqual([]);
  });

  it("reports not-configured when there are no sample rows", () => {
    const outcome = simulatePreview(config({ kind: "copy", sampleHeaders: HEADERS, sampleRows: [] }));
    expect(outcome.status).toBe("not-configured");
    expect(outcome.message).toMatch(/sample row/i);
  });

  it("reports not-configured when there are no sample headers", () => {
    const outcome = simulatePreview(config({ kind: "copy", sampleHeaders: [], sampleRows: [["x"]] }));
    expect(outcome.status).toBe("not-configured");
    expect(outcome.message).toMatch(/column headers/i);
  });

  it("reports an error, without throwing, when the key column is missing", () => {
    const outcome = simulatePreview(
      config({ kind: "aggregate", keyColumn: "Territory", valueColumn: "Amount", sampleHeaders: HEADERS, sampleRows: ROWS })
    );
    expect(outcome.status).toBe("error");
    expect(outcome.message).toContain("Territory");
    expect(outcome.after.rows).toEqual([]);
  });

  it("reports an error when the value column is missing", () => {
    const outcome = simulatePreview(
      config({ kind: "aggregate", keyColumn: "Region", valueColumn: "Total", sampleHeaders: HEADERS, sampleRows: ROWS })
    );
    expect(outcome.status).toBe("error");
    expect(outcome.message).toContain("Total");
  });
});

describe("simulatePreview — copy", () => {
  it("passes rows through unchanged", () => {
    const outcome = simulatePreview(config({ kind: "copy", sampleHeaders: HEADERS, sampleRows: ROWS }));
    expect(outcome.status).toBe("ok");
    expect(outcome.after.headers).toEqual(HEADERS);
    expect(outcome.after.rows).toEqual(ROWS);
    expect(outcome.before.rows).toEqual(ROWS);
  });
});

describe("simulatePreview — filter", () => {
  const base = { kind: "filter" as const, keyColumn: "Region", sampleHeaders: HEADERS, sampleRows: ROWS };

  it("equals", () => {
    const outcome = simulatePreview(config({ ...base, filterOperator: "equals", filterValue: "East" }));
    expect(outcome.after.rows).toEqual([
      ["East", "100"],
      ["East", "20"],
    ]);
  });

  it("not-equals", () => {
    const outcome = simulatePreview(config({ ...base, filterOperator: "not-equals", filterValue: "East" }));
    expect(outcome.after.rows.map((r) => r[0])).toEqual(["West", "West"]);
  });

  it("contains", () => {
    const outcome = simulatePreview(config({ ...base, filterOperator: "contains", filterValue: "es" }));
    expect(outcome.after.rows.map((r) => r[0])).toEqual(["West", "West"]);
  });

  it("greater-than and less-than on numbers", () => {
    const numeric = { ...base, keyColumn: "Amount" };
    const gt = simulatePreview(config({ ...numeric, filterOperator: "greater-than", filterValue: "60" }));
    expect(gt.after.rows.map((r) => r[1])).toEqual(["100", "70"]);
    const lt = simulatePreview(config({ ...numeric, filterOperator: "less-than", filterValue: "60" }));
    expect(lt.after.rows.map((r) => r[1])).toEqual(["50", "20"]);
  });

  it("is-blank and is-not-blank", () => {
    const rows = [
      ["East", "100"],
      ["", "50"],
      ["West", "20"],
    ];
    const blank = simulatePreview(
      config({ ...base, sampleRows: rows, filterOperator: "is-blank", filterValue: "ignored" })
    );
    expect(blank.after.rows).toEqual([["", "50"]]);
    const notBlank = simulatePreview(config({ ...base, sampleRows: rows, filterOperator: "is-not-blank" }));
    expect(notBlank.after.rows.map((r) => r[0])).toEqual(["East", "West"]);
  });

  it("notes when nothing matched", () => {
    const outcome = simulatePreview(config({ ...base, filterOperator: "equals", filterValue: "North" }));
    expect(outcome.status).toBe("ok");
    expect(outcome.after.rows).toEqual([]);
    expect(outcome.notes.join(" ")).toMatch(/no sample row matched/i);
  });
});

describe("simulatePreview — aggregate", () => {
  const base = {
    kind: "aggregate" as const,
    keyColumn: "Region",
    valueColumn: "Amount",
    sampleHeaders: HEADERS,
    sampleRows: ROWS,
  };

  it("sums by group and sorts by key", () => {
    const outcome = simulatePreview(config({ ...base, aggregate: "sum" }));
    expect(outcome.status).toBe("ok");
    expect(outcome.after.headers).toEqual(["Region", "Sum of Amount"]);
    expect(outcome.after.rows).toEqual([
      ["East", "120"],
      ["West", "120"],
    ]);
  });

  it("counts rows per group and ignores the value column", () => {
    const outcome = simulatePreview(config({ ...base, aggregate: "count", valueColumn: "" }));
    expect(outcome.after.headers).toEqual(["Region", "Count of rows"]);
    expect(outcome.after.rows).toEqual([
      ["East", "2"],
      ["West", "2"],
    ]);
  });

  it("averages, mins and maxes by group", () => {
    expect(simulatePreview(config({ ...base, aggregate: "average" })).after.rows).toEqual([
      ["East", "60"],
      ["West", "60"],
    ]);
    expect(simulatePreview(config({ ...base, aggregate: "min" })).after.rows).toEqual([
      ["East", "20"],
      ["West", "50"],
    ]);
    expect(simulatePreview(config({ ...base, aggregate: "max" })).after.rows).toEqual([
      ["East", "100"],
      ["West", "70"],
    ]);
  });

  it("parses currency and comma formatted numbers", () => {
    const outcome = simulatePreview(
      config({
        ...base,
        aggregate: "sum",
        sampleRows: [
          ["East", "$1,234.50"],
          ["East", "$0.50"],
        ],
      })
    );
    expect(outcome.after.rows).toEqual([["East", "1235"]]);
    expect(outcome.notes).toEqual([]);
  });

  it("skips a non-numeric value and reports it instead of treating it as zero", () => {
    const outcome = simulatePreview(
      config({
        ...base,
        aggregate: "sum",
        sampleRows: [
          ["East", "100"],
          ["East", "pending"],
        ],
      })
    );
    // 100, not 100 + 0 silently -- and definitely not an average over 2 rows.
    expect(outcome.after.rows).toEqual([["East", "100"]]);
    expect(outcome.notes.join(" ")).toContain("pending");
    expect(outcome.notes.join(" ")).toMatch(/not numeric/i);
  });

  it("does not average a skipped value into the denominator", () => {
    const outcome = simulatePreview(
      config({
        ...base,
        aggregate: "average",
        sampleRows: [
          ["East", "100"],
          ["East", "junk"],
        ],
      })
    );
    expect(outcome.after.rows).toEqual([["East", "100"]]);
  });

  it("shows a dash when a group has no numeric values at all", () => {
    const outcome = simulatePreview(
      config({ ...base, aggregate: "sum", sampleRows: [["East", "n/a"]] })
    );
    expect(outcome.after.rows).toEqual([["East", "—"]]);
    expect(outcome.notes.join(" ")).toMatch(/no numeric values/i);
  });
});

describe("simulatePreview — deduplicate", () => {
  it("keeps the first row per distinct key", () => {
    const outcome = simulatePreview(
      config({ kind: "deduplicate", keyColumn: "Region", sampleHeaders: HEADERS, sampleRows: ROWS })
    );
    expect(outcome.status).toBe("ok");
    expect(outcome.after.rows).toEqual([
      ["East", "100"],
      ["West", "50"],
    ]);
  });

  it("says the real macro's duplicate rule may differ", () => {
    const outcome = simulatePreview(
      config({ kind: "deduplicate", keyColumn: "Region", sampleHeaders: HEADERS, sampleRows: ROWS })
    );
    expect(outcome.steps.join(" ")).toMatch(/duplicate handling/i);
  });
});

describe("simulatePreview — caps", () => {
  it("truncates and notes excess rows", () => {
    const many = Array.from({ length: MAX_PREVIEW_ROWS + 10 }, (_, i) => ["East", String(i)]);
    const outcome = simulatePreview(config({ kind: "copy", sampleHeaders: HEADERS, sampleRows: many }));
    expect(outcome.after.rows.length).toBe(MAX_PREVIEW_ROWS);
    expect(outcome.notes.join(" ")).toContain(String(MAX_PREVIEW_ROWS));
  });

  it("truncates and notes excess columns", () => {
    const headers = Array.from({ length: MAX_PREVIEW_COLUMNS + 5 }, (_, i) => `C${i}`);
    const outcome = simulatePreview(
      config({ kind: "copy", sampleHeaders: headers, sampleRows: [headers.map(() => "x")] })
    );
    expect(outcome.after.headers.length).toBe(MAX_PREVIEW_COLUMNS);
    expect(outcome.notes.join(" ")).toContain(String(MAX_PREVIEW_COLUMNS));
  });

  it("caps rows parsed out of step 2 free text", () => {
    const text = Array.from({ length: MAX_PREVIEW_ROWS + 5 }, (_, i) => `East, ${i}`).join("\n");
    expect(parseSampleData("Region, Amount", text).rows.length).toBe(MAX_PREVIEW_ROWS);
  });
});
