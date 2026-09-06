// Validates a string as a usable VBA Sub/identifier name.

const VBA_RESERVED_WORDS = new Set(
  [
    "And", "As", "Boolean", "ByRef", "Byte", "ByVal", "Call", "Case", "Class",
    "Const", "Currency", "Debug", "Dim", "Do", "Double", "Each", "Else",
    "ElseIf", "Empty", "End", "Enum", "Eqv", "Erase", "Error", "Event",
    "Exit", "False", "For", "Friend", "Function", "Get", "Global", "GoSub",
    "GoTo", "If", "Imp", "Implements", "In", "Input", "Integer", "Is", "Let",
    "Lib", "Like", "Long", "Loop", "LSet", "Me", "Mod", "New", "Next", "Not",
    "Nothing", "Null", "Object", "On", "Option", "Optional", "Or", "ParamArray",
    "Preserve", "Print", "Private", "Property", "Public", "RaiseEvent",
    "ReDim", "Rem", "Resume", "RSet", "Select", "Set", "Single", "Static",
    "Stop", "String", "Sub", "Then", "To", "True", "Type", "TypeOf", "Until",
    "Variant", "Wend", "While", "With", "WithEvents", "Xor",
  ].map((w) => w.toLowerCase())
);

export const VBA_NAME_MIN_LENGTH = 1;
export const VBA_NAME_MAX_LENGTH = 40; // conservative; real VBA limit is 255

export interface VbaNameValidation {
  valid: boolean;
  errors: string[];
}

/**
 * Validates that `name` is usable as a VBA Sub name / identifier:
 * letters, digits, underscore; must start with a letter; not a reserved
 * word; within a reasonable length.
 */
export function validateVbaMacroName(name: string): VbaNameValidation {
  const errors: string[] = [];
  const trimmed = name.trim();

  if (trimmed.length === 0) {
    errors.push("Macro name is required.");
    return { valid: false, errors };
  }

  if (trimmed !== name) {
    errors.push("Macro name cannot have leading or trailing spaces.");
  }

  if (trimmed.length > VBA_NAME_MAX_LENGTH) {
    errors.push(`Macro name must be ${VBA_NAME_MAX_LENGTH} characters or fewer.`);
  }

  if (!/^[A-Za-z]/.test(trimmed)) {
    errors.push("Macro name must start with a letter.");
  }

  if (!/^[A-Za-z0-9_]*$/.test(trimmed)) {
    errors.push("Macro name can only contain letters, digits, and underscores (no spaces or punctuation).");
  }

  if (VBA_RESERVED_WORDS.has(trimmed.toLowerCase())) {
    errors.push(`"${trimmed}" is a reserved VBA keyword and cannot be used as a macro name.`);
  }

  return { valid: errors.length === 0, errors };
}
