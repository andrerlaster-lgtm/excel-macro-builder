import { emptyFormData, emptyPreviewConfig, MacroFormData, PreviewConfig, SavedProject } from "./types";

// Bump this and add a migration step in `migrateProject` whenever the saved
// project shape changes. No permanent delete UI exists in this phase --
// archiving is the only removal path; use browser site-data controls to
// actually clear storage (documented in the README).

export const CURRENT_SCHEMA_VERSION = 2;
export const STORAGE_KEY = "excel-macro-builder:projects:v1";

interface StorageEnvelope {
  schemaVersion: number;
  projects: SavedProject[];
}

function safeRandomId(): string {
  if (typeof crypto !== "undefined" && "randomUUID" in crypto) {
    return crypto.randomUUID();
  }
  return `id-${Date.now()}-${Math.random().toString(36).slice(2)}`;
}

function isPlainObject(v: unknown): v is Record<string, unknown> {
  return typeof v === "object" && v !== null && !Array.isArray(v);
}

/**
 * v1 -> v2: `form.preview` did not exist. Backfill it (and repair a
 * partially-written one) so older drafts open with a usable, unconfigured
 * preview instead of crashing the review step.
 */
function migrateFormToV2(form: MacroFormData): MacroFormData {
  const raw = (form as unknown as Record<string, unknown>).preview;
  if (!isPlainObject(raw)) {
    return { ...form, preview: emptyPreviewConfig() };
  }
  const defaults = emptyPreviewConfig();
  const preview: PreviewConfig = {
    kind: typeof raw.kind === "string" ? (raw.kind as PreviewConfig["kind"]) : defaults.kind,
    keyColumn: typeof raw.keyColumn === "string" ? raw.keyColumn : defaults.keyColumn,
    valueColumn: typeof raw.valueColumn === "string" ? raw.valueColumn : defaults.valueColumn,
    aggregate: typeof raw.aggregate === "string" ? (raw.aggregate as PreviewConfig["aggregate"]) : defaults.aggregate,
    filterOperator:
      typeof raw.filterOperator === "string"
        ? (raw.filterOperator as PreviewConfig["filterOperator"])
        : defaults.filterOperator,
    filterValue: typeof raw.filterValue === "string" ? raw.filterValue : defaults.filterValue,
    sampleHeaders: Array.isArray(raw.sampleHeaders) ? raw.sampleHeaders.map((h) => String(h ?? "")) : [],
    sampleRows: Array.isArray(raw.sampleRows)
      ? raw.sampleRows.map((row) => (Array.isArray(row) ? row.map((c) => String(c ?? "")) : []))
      : [],
  };
  return { ...form, preview };
}

/**
 * Migrates a single project record forward to the current schema version.
 * Unknown/missing fields are backfilled with safe defaults so older or
 * partially-written records never crash the app.
 */
export function migrateProject(raw: unknown): SavedProject | null {
  if (!isPlainObject(raw)) return null;

  const version = typeof raw.schemaVersion === "number" ? raw.schemaVersion : 0;

  // version 0 -> 1: introduce schemaVersion, archived flag, lastSpecification/lastResult.
  const base: SavedProject = {
    id: typeof raw.id === "string" && raw.id.length > 0 ? raw.id : safeRandomId(),
    schemaVersion: CURRENT_SCHEMA_VERSION,
    title: typeof raw.title === "string" ? raw.title : "Untitled macro project",
    createdAt: typeof raw.createdAt === "string" ? raw.createdAt : new Date().toISOString(),
    updatedAt: typeof raw.updatedAt === "string" ? raw.updatedAt : new Date().toISOString(),
    archived: typeof raw.archived === "boolean" ? raw.archived : false,
    form: migrateFormToV2(isPlainObject(raw.form) ? (raw.form as unknown as MacroFormData) : emptyFormData()),
    lastSpecification: isPlainObject(raw.lastSpecification)
      ? (raw.lastSpecification as unknown as SavedProject["lastSpecification"])
      : null,
    lastResult: isPlainObject(raw.lastResult) ? (raw.lastResult as unknown as SavedProject["lastResult"]) : null,
  };

  // version 1 -> 2: `form.preview` added (handled by migrateFormToV2 above,
  // which is safe to run on records that already carry it).
  void version;
  return base;
}

function readEnvelope(): StorageEnvelope {
  if (typeof window === "undefined") {
    return { schemaVersion: CURRENT_SCHEMA_VERSION, projects: [] };
  }
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (!raw) return { schemaVersion: CURRENT_SCHEMA_VERSION, projects: [] };
    const parsed = JSON.parse(raw);
    const list = isPlainObject(parsed) && Array.isArray(parsed.projects) ? parsed.projects : [];
    const migrated = list.map(migrateProject).filter((p): p is SavedProject => p !== null);
    return { schemaVersion: CURRENT_SCHEMA_VERSION, projects: migrated };
  } catch {
    return { schemaVersion: CURRENT_SCHEMA_VERSION, projects: [] };
  }
}

function writeEnvelope(envelope: StorageEnvelope): void {
  if (typeof window === "undefined") return;
  window.localStorage.setItem(STORAGE_KEY, JSON.stringify(envelope));
}

export function listProjects(): SavedProject[] {
  return readEnvelope().projects.sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
}

export function getProject(id: string): SavedProject | undefined {
  return readEnvelope().projects.find((p) => p.id === id);
}

export function createProject(title: string, form: MacroFormData = emptyFormData()): SavedProject {
  const now = new Date().toISOString();
  const project: SavedProject = {
    id: safeRandomId(),
    schemaVersion: CURRENT_SCHEMA_VERSION,
    title: title.trim() || "Untitled macro project",
    createdAt: now,
    updatedAt: now,
    archived: false,
    form,
    lastSpecification: null,
    lastResult: null,
  };
  const envelope = readEnvelope();
  envelope.projects.push(project);
  writeEnvelope(envelope);
  return project;
}

export function saveProject(project: SavedProject): void {
  const envelope = readEnvelope();
  const idx = envelope.projects.findIndex((p) => p.id === project.id);
  const updated = { ...project, updatedAt: new Date().toISOString() };
  if (idx === -1) {
    envelope.projects.push(updated);
  } else {
    envelope.projects[idx] = updated;
  }
  writeEnvelope(envelope);
}

export function renameProject(id: string, title: string): SavedProject | undefined {
  const envelope = readEnvelope();
  const project = envelope.projects.find((p) => p.id === id);
  if (!project) return undefined;
  project.title = title.trim() || project.title;
  project.updatedAt = new Date().toISOString();
  writeEnvelope(envelope);
  return project;
}

export function duplicateProject(id: string): SavedProject | undefined {
  const envelope = readEnvelope();
  const source = envelope.projects.find((p) => p.id === id);
  if (!source) return undefined;
  const now = new Date().toISOString();
  const copy: SavedProject = {
    ...source,
    id: safeRandomId(),
    title: `${source.title} (copy)`,
    createdAt: now,
    updatedAt: now,
    archived: false,
  };
  envelope.projects.push(copy);
  writeEnvelope(envelope);
  return copy;
}

export function setArchived(id: string, archived: boolean): SavedProject | undefined {
  const envelope = readEnvelope();
  const project = envelope.projects.find((p) => p.id === id);
  if (!project) return undefined;
  project.archived = archived;
  project.updatedAt = new Date().toISOString();
  writeEnvelope(envelope);
  return project;
}
