"use client";

import { ExcelPlatform, TaskDescription } from "@/lib/types";
import { FieldError } from "@/lib/formValidation";
import { TextField } from "./FormFields";

function errFor(errors: FieldError[], field: string): string | undefined {
  return errors.find((e) => e.field === field)?.message;
}

const PLATFORM_OPTIONS: { value: ExcelPlatform; label: string }[] = [
  { value: "windows", label: "Windows Excel" },
  { value: "mac", label: "Mac Excel" },
  { value: "cross-platform", label: "Best-effort cross-platform" },
];

export function DescribeStep({
  task,
  onChange,
  errors,
}: {
  task: TaskDescription;
  onChange: (task: TaskDescription) => void;
  errors: FieldError[];
}) {
  return (
    <div className="card">
      <h2 className="card-title">1. Describe the task</h2>
      <p className="card-subtitle">
        Plain language is fine here. Use only fictional example names if you want to illustrate the problem —
        never real client, company, or workbook data.
      </p>

      <TextField
        id="projectTitle"
        label="Project title"
        value={task.projectTitle}
        onChange={(v) => onChange({ ...task, projectTitle: v })}
        placeholder="e.g. Monthly Sales Rollup"
        error={errFor(errors, "task.projectTitle")}
        maxLength={200}
      />

      <TextField
        id="macroName"
        label="VBA macro name"
        value={task.macroName}
        onChange={(v) => onChange({ ...task, macroName: v })}
        placeholder="e.g. RollUpMonthlySales"
        hint="Letters, digits, underscore only; must start with a letter; not a reserved VBA word."
        error={errFor(errors, "task.macroName")}
        maxLength={40}
      />

      <fieldset>
        <legend>Target Excel environment</legend>
        <div className="radio-group" role="radiogroup" aria-label="Target Excel environment">
          {PLATFORM_OPTIONS.map((opt) => (
            <label className="radio-option" key={opt.value}>
              <input
                type="radio"
                name="platform"
                value={opt.value}
                checked={task.platform === opt.value}
                onChange={() => onChange({ ...task, platform: opt.value })}
              />
              {opt.label}
            </label>
          ))}
        </div>
      </fieldset>

      <TextField
        id="problemDescription"
        label="Plain-language description of the issue to automate"
        value={task.problemDescription}
        onChange={(v) => onChange({ ...task, problemDescription: v })}
        placeholder="e.g. Every month I manually copy new rows from a raw export into a summary sheet."
        error={errFor(errors, "task.problemDescription")}
        multiline
        maxLength={4000}
      />

      <TextField
        id="desiredResult"
        label="Desired final result"
        value={task.desiredResult}
        onChange={(v) => onChange({ ...task, desiredResult: v })}
        placeholder="e.g. The summary sheet has one row per region with totals for the month."
        error={errFor(errors, "task.desiredResult")}
        multiline
        maxLength={4000}
      />

      <TextField
        id="successCondition"
        label="Success condition — how you'll know it worked"
        value={task.successCondition}
        onChange={(v) => onChange({ ...task, successCondition: v })}
        placeholder="e.g. Region totals match the sum of the raw data, with no duplicate rows."
        error={errFor(errors, "task.successCondition")}
        multiline
        maxLength={4000}
      />
    </div>
  );
}
