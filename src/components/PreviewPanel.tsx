"use client";

import { useEffect, useState } from "react";
import { DataMapping, PreviewConfig } from "@/lib/types";
import { parseSampleData, PreviewTable, simulatePreview } from "@/lib/previewSimulator";
import { FieldWrapper } from "./FormFields";

function ResultTable({ title, table, caption }: { title: string; table: PreviewTable; caption: string }) {
  return (
    <div className="preview-table-block">
      <h4 className="preview-table-title">{title}</h4>
      <div className="table-scroll">
        <table className="preview-table">
          <caption className="visually-hidden">{caption}</caption>
          <thead>
            <tr>
              {table.headers.map((h, i) => (
                <th key={`${h}-${i}`} scope="col">
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {table.rows.length === 0 ? (
              <tr>
                <td colSpan={Math.max(1, table.headers.length)} className="preview-empty-cell">
                  No rows
                </td>
              </tr>
            ) : (
              table.rows.map((row, r) => (
                <tr key={r}>
                  {row.map((cell, c) => (
                    <td key={c}>{cell === "" ? <span className="preview-blank">(blank)</span> : cell}</td>
                  ))}
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
      <p className="preview-rowcount">
        {table.rows.length} row{table.rows.length === 1 ? "" : "s"}
      </p>
    </div>
  );
}

/**
 * Editable sample grid, reused for both the source sample and (for lookup
 * only) the destination sample. Each caller owns its own headers/rows and
 * its own add/remove-row wiring, so the two grids never interfere.
 */
function EditableSampleGrid({
  idPrefix,
  caption,
  emptyMessage,
  headers,
  rows,
  onSetCell,
  onAddRow,
  onRemoveRow,
}: {
  idPrefix: string;
  caption: string;
  emptyMessage: string;
  headers: string[];
  rows: string[][];
  onSetCell: (rowIndex: number, colIndex: number, value: string) => void;
  onAddRow: () => void;
  onRemoveRow: (rowIndex: number) => void;
}) {
  if (headers.length === 0) {
    return <p className="empty-state">{emptyMessage}</p>;
  }
  return (
    <>
      <div className="table-scroll">
        <table className="preview-table preview-table--editable">
          <caption className="visually-hidden">{caption}</caption>
          <thead>
            <tr>
              {headers.map((h, i) => (
                <th key={`${h}-${i}`} scope="col">
                  {h}
                </th>
              ))}
              <th scope="col">
                <span className="visually-hidden">Row actions</span>
              </th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row, r) => (
              <tr key={r}>
                {row.map((cell, c) => (
                  <td key={c}>
                    <label className="visually-hidden" htmlFor={`${idPrefix}-${r}-${c}`}>
                      {headers[c]}, row {r + 1}
                    </label>
                    <input
                      type="text"
                      id={`${idPrefix}-${r}-${c}`}
                      className="grid-input"
                      value={cell}
                      onChange={(e) => onSetCell(r, c, e.target.value)}
                    />
                  </td>
                ))}
                <td>
                  <button type="button" className="btn btn--danger-outline btn--small" onClick={() => onRemoveRow(r)}>
                    Remove<span className="visually-hidden"> row {r + 1}</span>
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="btn-row">
        <button type="button" className="btn" onClick={onAddRow}>
          Add row
        </button>
      </div>
    </>
  );
}

/**
 * Before/after preview of the operation the user picked in step 2.
 *
 * This runs entirely in the browser and needs no API key. It simulates the
 * DESCRIBED operation over the sample grid(s); it does not read, run, or
 * verify the generated VBA. The caution below says so, and must stay
 * visible.
 *
 * For every kind except "lookup" the sample grid is the SOURCE, and
 * before/after describe the source being transformed. Lookup is the
 * exception: it edits an EXISTING destination, so it gets a second grid for
 * the destination sample, and before/after describe the destination being
 * updated, not the source.
 */
export function PreviewPanel({
  mapping,
  preview,
  onPreviewChange,
}: {
  mapping: DataMapping;
  preview: PreviewConfig;
  onPreviewChange: (preview: PreviewConfig) => void;
}) {
  const outcome = simulatePreview(preview);
  const headers = preview.sampleHeaders;
  const rows = preview.sampleRows;
  const destHeaders = preview.destSampleHeaders;
  const destRows = preview.destSampleRows;
  const isLookup = preview.kind === "lookup";
  const hasSecondSource = isLookup && mapping.secondSourceEnabled;
  const source2Headers = preview.secondSourceSampleHeaders;
  const source2Rows = preview.secondSourceSampleRows;

  // Local text buffer for the destination headers field: committing only on
  // blur (not on every keystroke) means typing "Account, " does not collapse
  // back to "Account" mid-comma. Re-synced whenever the underlying headers
  // change from outside this input (e.g. a different project is loaded).
  const [destHeaderText, setDestHeaderText] = useState(destHeaders.join(", "));
  useEffect(() => {
    setDestHeaderText(destHeaders.join(", "));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [destHeaders.join(",")]);

  // Same pattern as the destination headers buffer above, for the second
  // source's grid.
  const [source2HeaderText, setSource2HeaderText] = useState(source2Headers.join(", "));
  useEffect(() => {
    setSource2HeaderText(source2Headers.join(", "));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [source2Headers.join(",")]);

  function setCell(rowIndex: number, colIndex: number, value: string) {
    const next = rows.map((row, r) => (r === rowIndex ? row.map((c, i) => (i === colIndex ? value : c)) : [...row]));
    onPreviewChange({ ...preview, sampleRows: next });
  }

  function addRow() {
    onPreviewChange({ ...preview, sampleRows: [...rows, headers.map(() => "")] });
  }

  function removeRow(rowIndex: number) {
    onPreviewChange({ ...preview, sampleRows: rows.filter((_, r) => r !== rowIndex) });
  }

  function setDestCell(rowIndex: number, colIndex: number, value: string) {
    const next = destRows.map((row, r) =>
      r === rowIndex ? row.map((c, i) => (i === colIndex ? value : c)) : [...row]
    );
    onPreviewChange({ ...preview, destSampleRows: next });
  }

  function addDestRow() {
    onPreviewChange({ ...preview, destSampleRows: [...destRows, destHeaders.map(() => "")] });
  }

  function removeDestRow(rowIndex: number) {
    onPreviewChange({ ...preview, destSampleRows: destRows.filter((_, r) => r !== rowIndex) });
  }

  function reseed() {
    const seeded = parseSampleData(
      mapping.sourceColumnHeaders,
      mapping.sourceExampleRows.notApplicable ? "" : mapping.sourceExampleRows.value
    );
    onPreviewChange({ ...preview, sampleHeaders: seeded.headers, sampleRows: seeded.rows });
  }

  /**
   * There is no step 2 field describing destination columns -- lookup is the
   * only operation that reads existing destination rows -- so its headers
   * are typed here directly, comma-separated, the same shape as step 2's own
   * "Relevant column headers" field. Existing row values are kept lined up
   * with the new header count (padded or trimmed), never dropped outright.
   */
  function setDestHeadersText(text: string) {
    const newHeaders = text
      .split(",")
      .map((h) => h.trim())
      .filter((h) => h.length > 0);
    const adjustedRows = destRows.map((row) => {
      const cells = row.slice(0, newHeaders.length);
      while (cells.length < newHeaders.length) cells.push("");
      return cells;
    });
    onPreviewChange({ ...preview, destSampleHeaders: newHeaders, destSampleRows: adjustedRows });
  }

  function setSource2Cell(rowIndex: number, colIndex: number, value: string) {
    const next = source2Rows.map((row, r) =>
      r === rowIndex ? row.map((c, i) => (i === colIndex ? value : c)) : [...row]
    );
    onPreviewChange({ ...preview, secondSourceSampleRows: next });
  }

  function addSource2Row() {
    onPreviewChange({ ...preview, secondSourceSampleRows: [...source2Rows, source2Headers.map(() => "")] });
  }

  function removeSource2Row(rowIndex: number) {
    onPreviewChange({ ...preview, secondSourceSampleRows: source2Rows.filter((_, r) => r !== rowIndex) });
  }

  /**
   * Same free-text header idiom as the destination grid, but the second
   * source DOES have a step 2 field for its headers (`secondSourceColumnHeaders`,
   * unlike the destination) -- so this also offers a one-click reseed from it.
   */
  function setSource2HeadersText(text: string) {
    const newHeaders = text
      .split(",")
      .map((h) => h.trim())
      .filter((h) => h.length > 0);
    const adjustedRows = source2Rows.map((row) => {
      const cells = row.slice(0, newHeaders.length);
      while (cells.length < newHeaders.length) cells.push("");
      return cells;
    });
    onPreviewChange({ ...preview, secondSourceSampleHeaders: newHeaders, secondSourceSampleRows: adjustedRows });
  }

  function reseedSource2Headers() {
    const newHeaders = mapping.secondSourceColumnHeaders
      .split(",")
      .map((h) => h.trim())
      .filter((h) => h.length > 0);
    const adjustedRows = source2Rows.map((row) => {
      const cells = row.slice(0, newHeaders.length);
      while (cells.length < newHeaders.length) cells.push("");
      return cells;
    });
    onPreviewChange({ ...preview, secondSourceSampleHeaders: newHeaders, secondSourceSampleRows: adjustedRows });
  }

  return (
    <section className="card" aria-labelledby="preview-heading">
      <h2 className="card-title" id="preview-heading">
        Simulated preview of the described operation
      </h2>

      <div className="callout callout--amber" role="note">
        <span className="callout-title">What this preview is, and is not</span>
        This simulates the operation you picked in step 2, applied to the sample rows below
        {isLookup
          ? hasSecondSource
            ? " (including the destination and second-source sample rows you type in below)"
            : " (including the destination sample rows you type in below)"
          : ""}
        . It does <strong>not</strong> run the generated VBA and cannot prove the real macro behaves this way — the
        macro can do something different. Read the generated code and test it on a copy of your workbook before
        trusting it.
      </div>

      <h3 className="section-heading">{isLookup ? "Source sample rows" : "Sample rows"}</h3>
      <p className="card-subtitle">
        Edit any cell to see the result change. These rows stay in your browser and are never sent anywhere.
      </p>

      <EditableSampleGrid
        idPrefix="sample"
        caption="Editable source sample rows used to compute the preview"
        emptyMessage='No sample columns yet. Enter your column headers in step 2, then use "Reseed from step 2".'
        headers={headers}
        rows={rows}
        onSetCell={setCell}
        onAddRow={addRow}
        onRemoveRow={removeRow}
      />

      <div className="btn-row">
        <button type="button" className="btn" onClick={reseed}>
          Reseed from step 2
        </button>
      </div>

      {isLookup && (
        <>
          <h3 className="section-heading">Destination sample rows</h3>
          <p className="card-subtitle">
            These represent EXISTING rows already in your destination report. There is no step 2 field to reseed
            them from, so set the column headers below, then type in a few rows yourself — include at least one row
            whose key has no match in the source sample above, to see how an unmatched row is left alone.
          </p>

          <FieldWrapper
            label="Destination column headers"
            htmlFor="dest-sample-headers-input"
            hint="Comma-separated, e.g. Account, debt, credit, ending."
          >
            <input
              type="text"
              id="dest-sample-headers-input"
              placeholder="e.g. Account, debt, credit, ending"
              value={destHeaderText}
              onChange={(e) => setDestHeaderText(e.target.value)}
              onBlur={(e) => setDestHeadersText(e.target.value)}
            />
          </FieldWrapper>

          <EditableSampleGrid
            idPrefix="dest-sample"
            caption="Editable destination sample rows used to compute the lookup preview"
            emptyMessage="No destination sample columns yet. Set the destination column headers above to start."
            headers={destHeaders}
            rows={destRows}
            onSetCell={setDestCell}
            onAddRow={addDestRow}
            onRemoveRow={removeDestRow}
          />
        </>
      )}

      {hasSecondSource && (
        <>
          <h3 className="section-heading">Second source sample rows</h3>
          <p className="card-subtitle">
            Matched against the destination by the same key as the source sample above. Include a row whose key does
            not appear in the source sample, to see a row matched only by the second source.
          </p>

          <FieldWrapper
            label="Second source column headers"
            htmlFor="source2-sample-headers-input"
            hint="Comma-separated, e.g. Account, Category, Note."
          >
            <input
              type="text"
              id="source2-sample-headers-input"
              placeholder="e.g. Account, Category, Note"
              value={source2HeaderText}
              onChange={(e) => setSource2HeaderText(e.target.value)}
              onBlur={(e) => setSource2HeadersText(e.target.value)}
            />
          </FieldWrapper>

          <div className="btn-row">
            <button type="button" className="btn" onClick={reseedSource2Headers}>
              Reseed headers from step 2
            </button>
          </div>

          <EditableSampleGrid
            idPrefix="source2-sample"
            caption="Editable second-source sample rows used to compute the lookup preview"
            emptyMessage="No second-source sample columns yet. Set the column headers above to start."
            headers={source2Headers}
            rows={source2Rows}
            onSetCell={setSource2Cell}
            onAddRow={addSource2Row}
            onRemoveRow={removeSource2Row}
          />
        </>
      )}

      <h3 className="section-heading">Before and after</h3>

      {outcome.status === "ok" ? (
        <>
          <div className="preview-grid">
            <ResultTable
              title="Before"
              table={outcome.before}
              caption={isLookup ? "Destination sample rows before the lookup" : "Sample rows before the operation"}
            />
            <ResultTable
              title="After"
              table={outcome.after}
              caption={isLookup ? "Destination sample rows after the lookup" : "Sample rows after the simulated operation"}
            />
          </div>

          {outcome.steps.length > 0 && (
            <>
              <h4 className="preview-table-title">What the simulation did</h4>
              <ol className="list-plain">
                {outcome.steps.map((s, i) => (
                  <li key={i}>{s}</li>
                ))}
              </ol>
            </>
          )}
        </>
      ) : (
        <p className="empty-state">{outcome.message}</p>
      )}

      {outcome.notes.length > 0 && (
        <div className="callout callout--amber" role="status" style={{ marginTop: 16 }}>
          <span className="callout-title">Notes on your sample data</span>
          <ul className="list-plain">
            {outcome.notes.map((n, i) => (
              <li key={i}>{n}</li>
            ))}
          </ul>
        </div>
      )}
    </section>
  );
}
