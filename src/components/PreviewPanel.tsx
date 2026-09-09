"use client";

import { DataMapping, PreviewConfig } from "@/lib/types";
import { parseSampleData, PreviewTable, simulatePreview } from "@/lib/previewSimulator";

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
 * Before/after preview of the operation the user picked in step 2.
 *
 * This runs entirely in the browser and needs no API key. It simulates the
 * DESCRIBED operation over the sample grid; it does not read, run, or verify
 * the generated VBA. The caution below says so, and must stay visible.
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

  function reseed() {
    const seeded = parseSampleData(
      mapping.sourceColumnHeaders,
      mapping.sourceExampleRows.notApplicable ? "" : mapping.sourceExampleRows.value
    );
    onPreviewChange({ ...preview, sampleHeaders: seeded.headers, sampleRows: seeded.rows });
  }

  return (
    <section className="card" aria-labelledby="preview-heading">
      <h2 className="card-title" id="preview-heading">
        Simulated preview of the described operation
      </h2>

      <div className="callout callout--amber" role="note">
        <span className="callout-title">What this preview is, and is not</span>
        This simulates the operation you picked in step 2, applied to the sample rows below. It does{" "}
        <strong>not</strong> run the generated VBA and cannot prove the real macro behaves this way — the macro can do
        something different. Read the generated code and test it on a copy of your workbook before trusting it.
      </div>

      <h3 className="section-heading">Sample rows</h3>
      <p className="card-subtitle">
        Edit any cell to see the result change. These rows stay in your browser and are never sent anywhere.
      </p>

      {headers.length === 0 ? (
        <p className="empty-state">
          No sample columns yet. Enter your column headers in step 2, then use &quot;Reseed from step 2&quot;.
        </p>
      ) : (
        <div className="table-scroll">
          <table className="preview-table preview-table--editable">
            <caption className="visually-hidden">Editable sample rows used to compute the preview</caption>
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
                      <label className="visually-hidden" htmlFor={`sample-${r}-${c}`}>
                        {headers[c]}, row {r + 1}
                      </label>
                      <input
                        type="text"
                        id={`sample-${r}-${c}`}
                        className="grid-input"
                        value={cell}
                        onChange={(e) => setCell(r, c, e.target.value)}
                      />
                    </td>
                  ))}
                  <td>
                    <button
                      type="button"
                      className="btn btn--danger-outline btn--small"
                      onClick={() => removeRow(r)}
                    >
                      Remove<span className="visually-hidden"> row {r + 1}</span>
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      <div className="btn-row">
        <button type="button" className="btn" onClick={addRow} disabled={headers.length === 0}>
          Add row
        </button>
        <button type="button" className="btn" onClick={reseed}>
          Reseed from step 2
        </button>
      </div>

      <h3 className="section-heading">Before and after</h3>

      {outcome.status === "ok" ? (
        <>
          <div className="preview-grid">
            <ResultTable title="Before" table={outcome.before} caption="Sample rows before the operation" />
            <ResultTable title="After" table={outcome.after} caption="Sample rows after the simulated operation" />
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
