"use client";

import {
  AggregateFn,
  AppendOrOverwrite,
  DataMapping,
  FilterOperator,
  PreviewConfig,
  PreviewOperationKind,
  RunConfig,
  TriggerMode,
} from "@/lib/types";
import { FieldError } from "@/lib/formValidation";
import { parseSampleData } from "@/lib/previewSimulator";
import { FieldWrapper, MaybeField, TextField } from "./FormFields";

function errFor(errors: FieldError[], field: string): string | undefined {
  return errors.find((e) => e.field === field)?.message;
}

const PREVIEW_KIND_OPTIONS: { value: PreviewOperationKind; label: string }[] = [
  { value: "not-configured", label: "No preview" },
  { value: "copy", label: "Copy rows as-is" },
  { value: "filter", label: "Filter rows" },
  { value: "deduplicate", label: "Remove duplicates" },
  { value: "aggregate", label: "Group and total" },
];

const AGGREGATE_OPTIONS: { value: AggregateFn; label: string }[] = [
  { value: "sum", label: "Sum" },
  { value: "count", label: "Count" },
  { value: "average", label: "Average" },
  { value: "min", label: "Min" },
  { value: "max", label: "Max" },
];

const FILTER_OPERATOR_OPTIONS: { value: FilterOperator; label: string }[] = [
  { value: "equals", label: "equals" },
  { value: "not-equals", label: "does not equal" },
  { value: "contains", label: "contains" },
  { value: "greater-than", label: "is greater than" },
  { value: "less-than", label: "is less than" },
  { value: "is-blank", label: "is blank" },
  { value: "is-not-blank", label: "is not blank" },
];

/**
 * A column chooser that offers the headers typed in step 2 as a dropdown, and
 * falls back to a free-text box when no headers have been entered yet.
 */
function ColumnField({
  id,
  label,
  value,
  columns,
  onChange,
  hint,
}: {
  id: string;
  label: string;
  value: string;
  columns: string[];
  onChange: (value: string) => void;
  hint?: string;
}) {
  if (columns.length === 0) {
    return <TextField id={id} label={label} value={value} onChange={onChange} hint={hint} placeholder="e.g. Region" />;
  }
  const matched = columns.find((c) => c.toLowerCase() === value.trim().toLowerCase()) ?? "";
  return (
    <FieldWrapper label={label} htmlFor={id} hint={hint}>
      <select id={id} value={matched} onChange={(e) => onChange(e.target.value)}>
        <option value="">Choose a column…</option>
        {columns.map((c) => (
          <option key={c} value={c}>
            {c}
          </option>
        ))}
      </select>
    </FieldWrapper>
  );
}

const TRIGGER_OPTIONS: { value: TriggerMode; label: string }[] = [
  { value: "manual", label: "Run manually" },
  { value: "button", label: "From a button" },
  { value: "workbook-open", label: "On workbook open" },
  { value: "sheet-change", label: "On sheet change" },
];

const APPEND_OPTIONS: { value: AppendOrOverwrite; label: string }[] = [
  { value: "append", label: "Append new rows" },
  { value: "overwrite", label: "Overwrite destination" },
  { value: "not-applicable", label: "Not applicable" },
];

export function MapDataStep({
  mapping,
  run,
  preview,
  onMappingChange,
  onRunChange,
  onPreviewChange,
  errors,
}: {
  mapping: DataMapping;
  run: RunConfig;
  preview: PreviewConfig;
  onMappingChange: (mapping: DataMapping) => void;
  onRunChange: (run: RunConfig) => void;
  onPreviewChange: (preview: PreviewConfig) => void;
  errors: FieldError[];
}) {
  const isEventDriven = run.trigger === "workbook-open" || run.trigger === "sheet-change";

  const seeded = parseSampleData(
    mapping.sourceColumnHeaders,
    mapping.sourceExampleRows.notApplicable ? "" : mapping.sourceExampleRows.value
  );

  function handlePreviewKind(kind: PreviewOperationKind) {
    if (kind === "not-configured") {
      onPreviewChange({ ...preview, kind });
      return;
    }
    // Seed the editable grid from step 2's free text the first time a
    // preview operation is chosen; never clobber an edited grid.
    const needsSeed = preview.sampleHeaders.length === 0;
    onPreviewChange({
      ...preview,
      kind,
      sampleHeaders: needsSeed ? seeded.headers : preview.sampleHeaders,
      sampleRows: needsSeed ? seeded.rows : preview.sampleRows,
    });
  }

  const previewColumns = preview.sampleHeaders.length > 0 ? preview.sampleHeaders : seeded.headers;
  const filterNeedsValue = preview.filterOperator !== "is-blank" && preview.filterOperator !== "is-not-blank";

  return (
    <div className="card">
      <h2 className="card-title">2. Map the Excel data</h2>
      <p className="card-subtitle">
        Workbook names/paths are generation context only — nothing here is opened, uploaded, or verified to exist.
        Use fictional examples, e.g. <code>Sample_Sales.xlsx</code>.
      </p>

      <h3 className="section-heading">Source</h3>
      <div className="field-row">
        <TextField
          id="sourceWorkbook"
          label="Source workbook file name/path"
          value={mapping.sourceWorkbook}
          onChange={(v) => onMappingChange({ ...mapping, sourceWorkbook: v })}
          placeholder="e.g. Sample_Sales.xlsx"
          error={errFor(errors, "mapping.sourceWorkbook")}
        />
        <TextField
          id="sourceWorksheet"
          label="Source worksheet"
          value={mapping.sourceWorksheet}
          onChange={(v) => onMappingChange({ ...mapping, sourceWorksheet: v })}
          placeholder="e.g. Raw Data"
          error={errFor(errors, "mapping.sourceWorksheet")}
        />
      </div>
      <div className="field-row">
        <TextField
          id="sourceRangeOrTable"
          label="Source range or table"
          value={mapping.sourceRangeOrTable}
          onChange={(v) => onMappingChange({ ...mapping, sourceRangeOrTable: v })}
          placeholder="e.g. A1:F500 or tblRawData"
          error={errFor(errors, "mapping.sourceRangeOrTable")}
        />
        <TextField
          id="sourceHeaderRow"
          label="Source header row"
          value={mapping.sourceHeaderRow}
          onChange={(v) => onMappingChange({ ...mapping, sourceHeaderRow: v })}
          placeholder="e.g. Row 1"
          error={errFor(errors, "mapping.sourceHeaderRow")}
        />
      </div>
      <TextField
        id="sourceColumnHeaders"
        label="Relevant column headers"
        value={mapping.sourceColumnHeaders}
        onChange={(v) => onMappingChange({ ...mapping, sourceColumnHeaders: v })}
        placeholder="e.g. Date, Region, Amount"
        multiline
        error={errFor(errors, "mapping.sourceColumnHeaders")}
      />
      <MaybeField
        id="sourceExampleRows"
        label="Example rows (optional)"
        value={mapping.sourceExampleRows}
        onChange={(v) => onMappingChange({ ...mapping, sourceExampleRows: v })}
        placeholder="e.g. 2026-01-05, East, 120.00"
      />

      <h3 className="section-heading">Destination</h3>
      <label className="radio-option" style={{ marginBottom: 12 }}>
        <input
          type="checkbox"
          checked={mapping.sameWorkbook}
          onChange={(e) => onMappingChange({ ...mapping, sameWorkbook: e.target.checked })}
        />
        Source and destination are the same workbook
      </label>

      {!mapping.sameWorkbook && (
        <MaybeField
          id="destinationWorkbook"
          label="Destination workbook file name/path"
          value={mapping.destinationWorkbook}
          onChange={(v) => onMappingChange({ ...mapping, destinationWorkbook: v })}
          placeholder="e.g. Monthly_Summary.xlsx"
          multiline={false}
          error={errFor(errors, "mapping.destinationWorkbook")}
        />
      )}

      <div className="field-row">
        <TextField
          id="destinationWorksheet"
          label="Destination worksheet"
          value={mapping.destinationWorksheet}
          onChange={(v) => onMappingChange({ ...mapping, destinationWorksheet: v })}
          placeholder="e.g. Monthly Summary"
          error={errFor(errors, "mapping.destinationWorksheet")}
        />
        <TextField
          id="destinationRangeOrTable"
          label="Destination table or starting cell"
          value={mapping.destinationRangeOrTable}
          onChange={(v) => onMappingChange({ ...mapping, destinationRangeOrTable: v })}
          placeholder="e.g. A1 or tblSummary"
          error={errFor(errors, "mapping.destinationRangeOrTable")}
        />
      </div>

      <h3 className="section-heading">Rules</h3>
      <MaybeField
        id="matchField"
        label="Match/lookup field"
        value={mapping.matchField}
        onChange={(v) => onMappingChange({ ...mapping, matchField: v })}
        placeholder="e.g. Region"
        multiline={false}
      />
      <MaybeField
        id="filters"
        label="Filters"
        value={mapping.filters}
        onChange={(v) => onMappingChange({ ...mapping, filters: v })}
        placeholder="e.g. Exclude rows where Status = Cancelled"
      />
      <MaybeField
        id="transformationRules"
        label="Transformation/calculation rules"
        value={mapping.transformationRules}
        onChange={(v) => onMappingChange({ ...mapping, transformationRules: v })}
        placeholder="e.g. Sum Amount by Region"
      />
      <MaybeField
        id="sortOrder"
        label="Sort order"
        value={mapping.sortOrder}
        onChange={(v) => onMappingChange({ ...mapping, sortOrder: v })}
        placeholder="e.g. Region ascending"
        multiline={false}
      />
      <MaybeField
        id="duplicateHandling"
        label="Duplicate handling"
        value={mapping.duplicateHandling}
        onChange={(v) => onMappingChange({ ...mapping, duplicateHandling: v })}
        placeholder="e.g. Keep the most recent row per Region"
      />
      <MaybeField
        id="blankOrErrorHandling"
        label="Blank/error handling"
        value={mapping.blankOrErrorHandling}
        onChange={(v) => onMappingChange({ ...mapping, blankOrErrorHandling: v })}
        placeholder="e.g. Skip rows with a blank Amount"
      />

      <fieldset>
        <legend>Append vs. overwrite behavior</legend>
        <div className="radio-group" role="radiogroup" aria-label="Append vs overwrite">
          {APPEND_OPTIONS.map((opt) => (
            <label className="radio-option" key={opt.value}>
              <input
                type="radio"
                name="appendOrOverwrite"
                value={opt.value}
                checked={mapping.appendOrOverwrite === opt.value}
                onChange={() => onMappingChange({ ...mapping, appendOrOverwrite: opt.value })}
              />
              {opt.label}
            </label>
          ))}
        </div>
      </fieldset>

      <h3 className="section-heading">Preview operation — optional</h3>
      <p className="card-subtitle">
        Pick the closest match to what you described and step 3 will compute a before/after example from your sample
        rows. This is optional and it does not change the generated VBA — the free-text rules above are still the
        source of truth for generation.
      </p>

      <fieldset>
        <legend>Which operation should the preview simulate?</legend>
        <div className="radio-group" role="radiogroup" aria-label="Preview operation">
          {PREVIEW_KIND_OPTIONS.map((opt) => (
            <label className="radio-option" key={opt.value}>
              <input
                type="radio"
                name="previewKind"
                value={opt.value}
                checked={preview.kind === opt.value}
                onChange={() => handlePreviewKind(opt.value)}
              />
              {opt.label}
            </label>
          ))}
        </div>
      </fieldset>

      {(preview.kind === "filter" || preview.kind === "deduplicate" || preview.kind === "aggregate") && (
        <ColumnField
          id="previewKeyColumn"
          label={preview.kind === "filter" ? "Column to test" : "Group/key column"}
          value={preview.keyColumn}
          columns={previewColumns}
          onChange={(v) => onPreviewChange({ ...preview, keyColumn: v })}
          hint={
            previewColumns.length === 0
              ? "Enter your column headers above to pick from a list instead."
              : undefined
          }
        />
      )}

      {preview.kind === "filter" && (
        <div className="field-row">
          <FieldWrapper label="Condition" htmlFor="previewFilterOperator">
            <select
              id="previewFilterOperator"
              value={preview.filterOperator}
              onChange={(e) => onPreviewChange({ ...preview, filterOperator: e.target.value as FilterOperator })}
            >
              {FILTER_OPERATOR_OPTIONS.map((opt) => (
                <option key={opt.value} value={opt.value}>
                  {opt.label}
                </option>
              ))}
            </select>
          </FieldWrapper>
          {filterNeedsValue && (
            <TextField
              id="previewFilterValue"
              label="Compared with"
              value={preview.filterValue}
              onChange={(v) => onPreviewChange({ ...preview, filterValue: v })}
              placeholder="e.g. East"
            />
          )}
        </div>
      )}

      {preview.kind === "aggregate" && (
        <div className="field-row">
          <FieldWrapper label="Calculation" htmlFor="previewAggregate">
            <select
              id="previewAggregate"
              value={preview.aggregate}
              onChange={(e) => onPreviewChange({ ...preview, aggregate: e.target.value as AggregateFn })}
            >
              {AGGREGATE_OPTIONS.map((opt) => (
                <option key={opt.value} value={opt.value}>
                  {opt.label}
                </option>
              ))}
            </select>
          </FieldWrapper>
          {preview.aggregate !== "count" && (
            <ColumnField
              id="previewValueColumn"
              label="Value column (numeric)"
              value={preview.valueColumn}
              columns={previewColumns}
              onChange={(v) => onPreviewChange({ ...preview, valueColumn: v })}
            />
          )}
        </div>
      )}

      <h3 className="section-heading">How the macro is started</h3>
      <fieldset>
        <legend>Trigger</legend>
        <div className="radio-group" role="radiogroup" aria-label="Macro trigger">
          {TRIGGER_OPTIONS.map((opt) => (
            <label className="radio-option" key={opt.value}>
              <input
                type="radio"
                name="trigger"
                value={opt.value}
                checked={run.trigger === opt.value}
                onChange={() => onRunChange({ ...run, trigger: opt.value })}
              />
              {opt.label}
            </label>
          ))}
        </div>
      </fieldset>

      {isEventDriven && (
        <div className="callout callout--amber" role="alert" style={{ marginBottom: 16 }}>
          <span className="callout-title">Event-driven macros run automatically</span>
          A macro triggered by opening the workbook or changing a sheet runs without you clicking anything. It can
          fire more often than expected (e.g. on every edit) and cause unintended side effects. Review it especially
          carefully and test on a copy first.
        </div>
      )}

      <MaybeField
        id="constraints"
        label="Optional constraints"
        value={run.constraints}
        onChange={(v) => onRunChange({ ...run, constraints: v })}
        placeholder="e.g. Must finish in under 5 seconds, must not prompt the user"
      />
      <MaybeField
        id="mustNotChange"
        label="Anything the macro must not change"
        value={run.mustNotChange}
        onChange={(v) => onRunChange({ ...run, mustNotChange: v })}
        placeholder="e.g. Do not modify the Raw Data sheet"
      />
    </div>
  );
}
