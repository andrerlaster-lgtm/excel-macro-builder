"use client";

import { useEffect, useMemo, useState } from "react";
import { emptyFormData, MacroFormData, MacroGenerationResult, MacroSpecification, SavedProject } from "@/lib/types";
import { validateMacroForm, FieldError } from "@/lib/formValidation";
import { buildSpecification } from "@/lib/specBuilder";
import {
  createProject,
  duplicateProject,
  listProjects,
  renameProject,
  saveProject,
  setArchived,
} from "@/lib/storage";
import { Stepper, StepDef } from "./Stepper";
import { DescribeStep } from "./DescribeStep";
import { MapDataStep } from "./MapDataStep";
import { ReviewStep } from "./ReviewStep";
import { ResultPanel } from "./ResultPanel";
import { ProjectSidebar } from "./ProjectSidebar";

const STEPS: StepDef[] = [
  { key: "describe", label: "Describe" },
  { key: "map", label: "Map data" },
  { key: "review", label: "Review & generate" },
];

function isFormEmpty(form: MacroFormData): boolean {
  return JSON.stringify(form) === JSON.stringify(emptyFormData());
}

export function App() {
  const [projects, setProjects] = useState<SavedProject[]>([]);
  const [currentId, setCurrentId] = useState<string | null>(null);
  const [form, setForm] = useState<MacroFormData>(emptyFormData());
  const [stepIndex, setStepIndex] = useState(0);
  const [furthest, setFurthest] = useState(0);
  const [errors, setErrors] = useState<FieldError[]>([]);
  const [result, setResult] = useState<MacroGenerationResult | null>(null);
  const [specification, setSpecification] = useState<MacroSpecification | null>(null);
  const [generating, setGenerating] = useState(false);
  const [generationError, setGenerationError] = useState<string | null>(null);
  const [aiConfigured, setAiConfigured] = useState(true); // optimistic; corrected on first attempt/probe
  const [showArchived, setShowArchived] = useState(false);
  const [pendingAction, setPendingAction] = useState<null | { kind: "new" | "clear-generate" }>(null);

  useEffect(() => {
    setProjects(listProjects());
  }, []);

  function refreshProjects() {
    setProjects(listProjects());
  }

  function loadProject(id: string) {
    const project = projects.find((p) => p.id === id) ?? listProjects().find((p) => p.id === id);
    if (!project) return;
    setCurrentId(project.id);
    setForm(project.form);
    setResult(project.lastResult);
    setSpecification(project.lastSpecification);
    setStepIndex(0);
    setFurthest(project.lastResult ? 2 : 0);
    setErrors([]);
    setGenerationError(null);
  }

  function persistCurrent(updates: Partial<Pick<SavedProject, "form" | "lastResult" | "lastSpecification" | "title">>) {
    if (!currentId) return;
    const existing = listProjects().find((p) => p.id === currentId);
    if (!existing) return;
    const updated: SavedProject = {
      ...existing,
      ...updates,
    };
    saveProject(updated);
    refreshProjects();
  }

  function ensureProject(): string {
    if (currentId) return currentId;
    const title = form.task.projectTitle.trim() || "Untitled macro project";
    const project = createProject(title, form);
    setCurrentId(project.id);
    refreshProjects();
    return project.id;
  }

  function handleFormChange(next: MacroFormData) {
    setForm(next);
    const id = currentId;
    if (id) {
      persistCurrent({ form: next, title: next.task.projectTitle.trim() || undefined });
    }
  }

  function doStartNew() {
    setCurrentId(null);
    setForm(emptyFormData());
    setResult(null);
    setSpecification(null);
    setStepIndex(0);
    setFurthest(0);
    setErrors([]);
    setGenerationError(null);
  }

  function handleNewProject() {
    if (!isFormEmpty(form) || result) {
      setPendingAction({ kind: "new" });
      return;
    }
    doStartNew();
  }

  /** Errors relevant to a given step only, so step 0 isn't blocked by step 1's required fields. */
  function errorsForStep(allErrors: FieldError[], index: number): FieldError[] {
    if (index === 0) return allErrors.filter((e) => e.field.startsWith("task."));
    if (index === 1) return allErrors.filter((e) => e.field.startsWith("mapping.") || e.field.startsWith("run."));
    return allErrors;
  }

  function goToStep(index: number) {
    if (index === stepIndex) return;
    if (index > stepIndex) {
      const validation = validateMacroForm(form);
      // Check every step strictly between the current one and the target (inclusive of current).
      for (let s = stepIndex; s < index; s++) {
        if (errorsForStep(validation.errors, s).length > 0) {
          setErrors(validation.errors);
          return;
        }
      }
    }
    setStepIndex(index);
    setFurthest((f) => Math.max(f, index));
    ensureProject();
    persistCurrent({ form });
  }

  function handleNext() {
    const validation = validateMacroForm(form);
    const relevant = errorsForStep(validation.errors, stepIndex);
    setErrors(validation.errors);
    if (relevant.length > 0) return;
    ensureProject();
    persistCurrent({ form });
    const next = Math.min(stepIndex + 1, STEPS.length - 1);
    setStepIndex(next);
    setFurthest((f) => Math.max(f, next));
  }

  function handleBack() {
    setStepIndex((s) => Math.max(0, s - 1));
  }

  async function runGeneration() {
    setGenerating(true);
    setGenerationError(null);
    try {
      const res = await fetch("/api/generate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ form }),
      });
      const data = await res.json();

      if (res.status === 503 && data.code === "missing-api-key") {
        setAiConfigured(false);
        setGenerationError(data.error);
        return;
      }

      if (!res.ok || !data.ok) {
        setGenerationError(data.error || "Generation failed. Please try again.");
        if (data.specification) setSpecification(data.specification);
        return;
      }

      setAiConfigured(true);
      setSpecification(data.specification);
      setResult(data.result);
      const id = ensureProject();
      persistCurrent({
        form,
        lastSpecification: data.specification,
        lastResult: data.result,
        title: form.task.projectTitle.trim() || undefined,
      });
      void id;
    } catch {
      setGenerationError("Could not reach the server. Check your connection and try again.");
    } finally {
      setGenerating(false);
    }
  }

  function handleGenerateClick() {
    if (result) {
      setPendingAction({ kind: "clear-generate" });
      return;
    }
    void runGeneration();
  }

  function handleExportSpecOnly() {
    const spec = buildSpecification(form);
    setSpecification(spec);
    const blob = new Blob([JSON.stringify(spec, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${(spec.macroName || "macro-spec").replace(/[^A-Za-z0-9_-]+/g, "_")}.spec.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  const stepErrors = useMemo(() => errors, [errors]);

  return (
    <div className="app-shell">
      <header className="app-nav">
        <div className="app-nav__brand">
          <span className="app-nav__title">Excel Macro Builder</span>
          <span className="app-nav__subtitle">Draft reviewable VBA — never opens or runs your workbook</span>
        </div>
        <Stepper steps={STEPS} currentIndex={stepIndex} onSelect={goToStep} furthestReachable={furthest} />
      </header>

      <div className="app-body">
        <main className="app-main">
          {stepIndex === 0 && (
            <DescribeStep task={form.task} onChange={(task) => handleFormChange({ ...form, task })} errors={stepErrors} />
          )}
          {stepIndex === 1 && (
            <MapDataStep
              mapping={form.mapping}
              run={form.run}
              onMappingChange={(mapping) => handleFormChange({ ...form, mapping })}
              onRunChange={(run) => handleFormChange({ ...form, run })}
              errors={stepErrors}
            />
          )}
          {stepIndex === 2 && (
            <>
              <ReviewStep form={form} aiConfigured={aiConfigured} />
              {generationError && (
                <div className="callout callout--red" role="alert" style={{ marginTop: 16 }}>
                  <span className="callout-title">Generation problem</span>
                  {generationError}
                </div>
              )}
              <div className="btn-row">
                <button
                  type="button"
                  className="btn btn--primary"
                  onClick={handleGenerateClick}
                  disabled={generating}
                  aria-busy={generating}
                >
                  {generating && <span className="spinner" aria-hidden="true" />}
                  {generating ? "Generating…" : "Generate VBA"}
                </button>
                <button type="button" className="btn" onClick={handleExportSpecOnly}>
                  Download specification (.json)
                </button>
              </div>
            </>
          )}

          {result && specification && stepIndex === 2 && (
            <ResultPanel
              result={result}
              specification={specification}
              macroName={form.task.macroName}
              onStartOver={handleNewProject}
            />
          )}

          {stepIndex < 2 && (
            <div className="btn-row">
              {stepIndex > 0 && (
                <button type="button" className="btn" onClick={handleBack}>
                  Back
                </button>
              )}
              <button type="button" className="btn btn--primary" onClick={handleNext}>
                Continue
              </button>
            </div>
          )}
          {stepIndex === 2 && (
            <div className="btn-row">
              <button type="button" className="btn" onClick={handleBack}>
                Back
              </button>
            </div>
          )}
        </main>

        <aside className="app-sidebar">
          <ProjectSidebar
            projects={projects}
            currentId={currentId}
            onOpen={loadProject}
            onNew={handleNewProject}
            onRename={(id, title) => {
              renameProject(id, title);
              refreshProjects();
            }}
            onDuplicate={(id) => {
              duplicateProject(id);
              refreshProjects();
            }}
            onArchiveToggle={(id, archived) => {
              setArchived(id, archived);
              refreshProjects();
            }}
            showArchived={showArchived}
            onToggleShowArchived={() => setShowArchived((v) => !v)}
          />
        </aside>
      </div>

      {pendingAction && (
        <div
          role="dialog"
          aria-modal="true"
          aria-label="Confirm action"
          style={{
            position: "fixed",
            inset: 0,
            background: "rgba(16,38,30,0.45)",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            padding: 16,
          }}
        >
          <div className="card" style={{ maxWidth: 420 }}>
            <h2 className="card-title">
              {pendingAction.kind === "new" ? "Start a new project?" : "Replace the generated result?"}
            </h2>
            <p className="card-subtitle">
              {pendingAction.kind === "new"
                ? "This form has unsaved input. Starting a new project will clear it from the editor (it stays saved under the current project if one was created)."
                : "You already have generated VBA and notes. Generating again will replace them."}
            </p>
            <div className="btn-row">
              <button
                type="button"
                className="btn btn--primary"
                onClick={() => {
                  const kind = pendingAction.kind;
                  setPendingAction(null);
                  if (kind === "new") doStartNew();
                  else void runGeneration();
                }}
              >
                {pendingAction.kind === "new" ? "Start new" : "Replace and generate"}
              </button>
              <button type="button" className="btn" onClick={() => setPendingAction(null)}>
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
