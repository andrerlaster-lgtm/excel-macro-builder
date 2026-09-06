// Derives a safe .bas filename from a macro name, sanitized for Windows/Mac.

const WINDOWS_RESERVED_NAMES = new Set([
  "con", "prn", "aux", "nul",
  "com1", "com2", "com3", "com4", "com5", "com6", "com7", "com8", "com9",
  "lpt1", "lpt2", "lpt3", "lpt4", "lpt5", "lpt6", "lpt7", "lpt8", "lpt9",
]);

/**
 * Produces a filesystem-safe .bas filename from a macro name (or arbitrary
 * text). Falls back to "macro" if nothing usable remains.
 */
export function safeBasFilename(rawName: string): string {
  let base = (rawName || "").trim();

  // Strip characters invalid on Windows/Mac filesystems, keep ASCII-safe.
  base = base.replace(/[^A-Za-z0-9_-]+/g, "_");
  base = base.replace(/_{2,}/g, "_");
  base = base.replace(/^_+|_+$/g, "");

  if (base.length === 0) {
    base = "macro";
  }

  if (WINDOWS_RESERVED_NAMES.has(base.toLowerCase())) {
    base = `${base}_macro`;
  }

  // Keep filenames reasonable in length.
  if (base.length > 100) {
    base = base.slice(0, 100);
  }

  return `${base}.bas`;
}
