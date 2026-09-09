import { beforeEach, describe, expect, it } from "vitest";
import { emptyFormData, emptyPreviewConfig } from "./types";
import { simulatePreview } from "./previewSimulator";
import {
  CURRENT_SCHEMA_VERSION,
  createProject,
  duplicateProject,
  getProject,
  listProjects,
  migrateProject,
  renameProject,
  setArchived,
  STORAGE_KEY,
} from "./storage";

// jsdom is not configured for this project (node test environment), so we
// provide a minimal localStorage shim on globalThis for these tests.
class MemoryStorage {
  private store = new Map<string, string>();
  getItem(key: string) {
    return this.store.has(key) ? this.store.get(key)! : null;
  }
  setItem(key: string, value: string) {
    this.store.set(key, value);
  }
  removeItem(key: string) {
    this.store.delete(key);
  }
  clear() {
    this.store.clear();
  }
}

beforeEach(() => {
  (globalThis as unknown as { window: unknown }).window = {
    localStorage: new MemoryStorage(),
  };
});

describe("storage", () => {
  it("creates and lists a project", () => {
    const project = createProject("My Macro", emptyFormData());
    const list = listProjects();
    expect(list.some((p) => p.id === project.id)).toBe(true);
    expect(project.schemaVersion).toBe(CURRENT_SCHEMA_VERSION);
    expect(project.archived).toBe(false);
  });

  it("gets a project by id", () => {
    const project = createProject("Findable", emptyFormData());
    const found = getProject(project.id);
    expect(found?.title).toBe("Findable");
  });

  it("renames a project", () => {
    const project = createProject("Old Name", emptyFormData());
    renameProject(project.id, "New Name");
    expect(getProject(project.id)?.title).toBe("New Name");
  });

  it("duplicates a project with a new id", () => {
    const project = createProject("Original", emptyFormData());
    const copy = duplicateProject(project.id);
    expect(copy?.id).not.toBe(project.id);
    expect(copy?.title).toContain("copy");
    expect(listProjects().length).toBe(2);
  });

  it("archives and restores a project", () => {
    const project = createProject("Archivable", emptyFormData());
    setArchived(project.id, true);
    expect(getProject(project.id)?.archived).toBe(true);
    setArchived(project.id, false);
    expect(getProject(project.id)?.archived).toBe(false);
  });

  it("migrates a legacy record missing schemaVersion/archived fields", () => {
    const legacy = {
      id: "legacy-1",
      title: "Legacy Project",
      form: emptyFormData(),
    };
    const migrated = migrateProject(legacy);
    expect(migrated).not.toBeNull();
    expect(migrated?.schemaVersion).toBe(CURRENT_SCHEMA_VERSION);
    expect(migrated?.archived).toBe(false);
    expect(migrated?.lastResult).toBeNull();
  });

  it("migrates a v1 record forward to v2 with a usable default preview config", () => {
    const v1Form = emptyFormData() as Partial<ReturnType<typeof emptyFormData>>;
    delete v1Form.preview; // v1 forms had no preview section
    const v1Record = {
      id: "v1-project",
      schemaVersion: 1,
      title: "Pre-preview Draft",
      createdAt: "2026-01-01T00:00:00.000Z",
      updatedAt: "2026-01-01T00:00:00.000Z",
      archived: false,
      form: v1Form,
      lastSpecification: null,
      lastResult: null,
    };

    const migrated = migrateProject(v1Record);
    expect(migrated).not.toBeNull();
    expect(migrated?.schemaVersion).toBe(2);
    expect(migrated?.title).toBe("Pre-preview Draft");
    expect(migrated?.form.preview).toEqual(emptyPreviewConfig());
    // ...and it is immediately usable by the simulator without throwing.
    expect(simulatePreview(migrated!.form.preview).status).toBe("not-configured");
  });

  it("repairs a partially-written preview config instead of trusting it", () => {
    const migrated = migrateProject({
      id: "half-written",
      schemaVersion: 2,
      form: { ...emptyFormData(), preview: { kind: "copy", sampleRows: "not an array" } },
    });
    expect(migrated?.form.preview.kind).toBe("copy");
    expect(migrated?.form.preview.sampleRows).toEqual([]);
    expect(migrated?.form.preview.sampleHeaders).toEqual([]);
    expect(migrated?.form.preview.aggregate).toBe("sum");
  });

  it("drops unparseable records instead of crashing", () => {
    expect(migrateProject(null)).toBeNull();
    expect(migrateProject("not an object")).toBeNull();
    expect(migrateProject(42)).toBeNull();
  });

  it("survives a corrupted localStorage payload by returning an empty list", () => {
    (window as unknown as { localStorage: MemoryStorage }).localStorage.setItem(STORAGE_KEY, "{not json");
    expect(listProjects()).toEqual([]);
  });
});
