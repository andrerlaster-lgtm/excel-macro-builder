"use client";

import { Maybe } from "@/lib/types";

interface FieldWrapperProps {
  label: string;
  htmlFor: string;
  hint?: string;
  error?: string;
  children: React.ReactNode;
}

export function FieldWrapper({ label, htmlFor, hint, error, children }: FieldWrapperProps) {
  return (
    <div className="field-group">
      <label className="field-label" htmlFor={htmlFor}>
        {label}
      </label>
      {children}
      {hint && !error && <span className="field-hint">{hint}</span>}
      {error && (
        <span className="field-error" role="alert">
          {error}
        </span>
      )}
    </div>
  );
}

interface TextFieldProps {
  id: string;
  label: string;
  value: string;
  onChange: (value: string) => void;
  hint?: string;
  error?: string;
  placeholder?: string;
  maxLength?: number;
  multiline?: boolean;
}

export function TextField({
  id,
  label,
  value,
  onChange,
  hint,
  error,
  placeholder,
  maxLength,
  multiline,
}: TextFieldProps) {
  return (
    <FieldWrapper label={label} htmlFor={id} hint={hint} error={error}>
      {multiline ? (
        <textarea
          id={id}
          value={value}
          onChange={(e) => onChange(e.target.value)}
          placeholder={placeholder}
          maxLength={maxLength}
          aria-invalid={Boolean(error)}
        />
      ) : (
        <input
          type="text"
          id={id}
          value={value}
          onChange={(e) => onChange(e.target.value)}
          placeholder={placeholder}
          maxLength={maxLength}
          aria-invalid={Boolean(error)}
        />
      )}
    </FieldWrapper>
  );
}

interface MaybeFieldProps {
  id: string;
  label: string;
  value: Maybe;
  onChange: (value: Maybe) => void;
  hint?: string;
  error?: string;
  placeholder?: string;
  multiline?: boolean;
}

/**
 * A text field paired with an explicit "Not applicable" checkbox, so
 * optional concerns can be marked out-of-scope instead of forcing invented
 * details.
 */
export function MaybeField({ id, label, value, onChange, hint, error, placeholder, multiline = true }: MaybeFieldProps) {
  const naId = `${id}-na`;
  return (
    <FieldWrapper label={label} htmlFor={id} hint={hint} error={error}>
      {multiline ? (
        <textarea
          id={id}
          value={value.value}
          disabled={value.notApplicable}
          onChange={(e) => onChange({ notApplicable: false, value: e.target.value })}
          placeholder={placeholder}
          aria-invalid={Boolean(error)}
        />
      ) : (
        <input
          type="text"
          id={id}
          value={value.value}
          disabled={value.notApplicable}
          onChange={(e) => onChange({ notApplicable: false, value: e.target.value })}
          placeholder={placeholder}
          aria-invalid={Boolean(error)}
        />
      )}
      <label className="na-row" htmlFor={naId}>
        <input
          type="checkbox"
          id={naId}
          checked={value.notApplicable}
          onChange={(e) => onChange({ notApplicable: e.target.checked, value: value.value })}
        />
        Not applicable
      </label>
    </FieldWrapper>
  );
}
