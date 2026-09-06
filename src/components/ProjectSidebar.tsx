"use client";

import { useState } from "react";
import { SavedProject } from "@/lib/types";

export function ProjectSidebar({
  projects,
  currentId,
  onOpen,
  onNew,
  onRename,
  onDuplicate,
  onArchiveToggle,
  showArchived,
  onToggleShowArchived,
}: {
  projects: SavedProject[];
  currentId: string | null;
  onOpen: (id: string) => void;
  onNew: () => void;
  onRename: (id: string, title: string) => void;
  onDuplicate: (id: string) => void;
  onArchiveToggle: (id: string, archived: boolean) => void;
  showArchived: boolean;
  onToggleShowArchived: () => void;
}) {
  const [renamingId, setRenamingId] = useState<string | null>(null);
  const [draftTitle, setDraftTitle] = useState("");

  const visible = projects.filter((p) => showArchived || !p.archived);

  return (
    <div className="card" aria-label="Saved macro projects">
      <h2 className="card-title">Projects</h2>
      <p className="card-subtitle">Saved locally in this browser only. No sign-in, no cloud sync.</p>

      <button type="button" className="btn btn--primary" style={{ width: "100%", justifyContent: "center" }} onClick={onNew}>
        + New project
      </button>

      <label className="radio-option" style={{ marginTop: 12 }}>
        <input type="checkbox" checked={showArchived} onChange={onToggleShowArchived} />
        Show archived
      </label>

      {visible.length === 0 ? (
        <p className="empty-state">No projects yet. Start one above.</p>
      ) : (
        <ul className="project-list" style={{ marginTop: 12 }}>
          {visible.map((p) => (
            <li key={p.id} className="project-item" data-current={p.id === currentId}>
              {renamingId === p.id ? (
                <form
                  onSubmit={(e) => {
                    e.preventDefault();
                    onRename(p.id, draftTitle);
                    setRenamingId(null);
                  }}
                >
                  <input
                    type="text"
                    value={draftTitle}
                    onChange={(e) => setDraftTitle(e.target.value)}
                    aria-label="Rename project"
                    autoFocus
                  />
                  <div className="project-item__actions">
                    <button type="submit" className="btn btn--primary">
                      Save
                    </button>
                    <button type="button" className="btn" onClick={() => setRenamingId(null)}>
                      Cancel
                    </button>
                  </div>
                </form>
              ) : (
                <>
                  <button type="button" className="project-item__title" onClick={() => onOpen(p.id)}>
                    {p.title} {p.archived && <span className="badge">Archived</span>}
                  </button>
                  <div className="project-item__meta">Updated {new Date(p.updatedAt).toLocaleString()}</div>
                  <div className="project-item__actions">
                    <button
                      type="button"
                      className="btn"
                      onClick={() => {
                        setRenamingId(p.id);
                        setDraftTitle(p.title);
                      }}
                    >
                      Rename
                    </button>
                    <button type="button" className="btn" onClick={() => onDuplicate(p.id)}>
                      Duplicate
                    </button>
                    <button type="button" className="btn" onClick={() => onArchiveToggle(p.id, !p.archived)}>
                      {p.archived ? "Restore" : "Archive"}
                    </button>
                  </div>
                </>
              )}
            </li>
          ))}
        </ul>
      )}

      <p className="field-hint" style={{ marginTop: 12 }}>
        There is no permanent delete in this app. To fully remove data, clear this site&apos;s data in your browser
        settings (see README).
      </p>
    </div>
  );
}
