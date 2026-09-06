"use client";

export interface StepDef {
  key: string;
  label: string;
}

export function Stepper({
  steps,
  currentIndex,
  onSelect,
  furthestReachable,
}: {
  steps: StepDef[];
  currentIndex: number;
  onSelect: (index: number) => void;
  furthestReachable: number;
}) {
  return (
    <nav aria-label="Guided workflow steps">
      <ol className="stepper">
        {steps.map((step, index) => {
          const active = index === currentIndex;
          const complete = index < currentIndex;
          const reachable = index <= furthestReachable;
          return (
            <li key={step.key}>
              <button
                type="button"
                className="stepper__item"
                data-active={active}
                data-complete={complete}
                aria-current={active ? "step" : undefined}
                disabled={!reachable}
                onClick={() => reachable && onSelect(index)}
                style={{ cursor: reachable ? "pointer" : "not-allowed" }}
              >
                <span className="stepper__badge" aria-hidden="true">
                  {complete ? "✓" : index + 1}
                </span>
                {step.label}
              </button>
            </li>
          );
        })}
      </ol>
    </nav>
  );
}
