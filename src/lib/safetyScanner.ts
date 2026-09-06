// Static warning pass: scans generated VBA text for potentially destructive
// or external-facing patterns. This is a review aid ONLY -- it does not
// prove the code is safe, and its absence of findings is not a safety
// guarantee. The UI must always say so alongside any results.

export type WarningCategory =
  | "file-deletion"
  | "shell-execution"
  | "bulk-row-or-column-deletion"
  | "workbook-overwrite"
  | "email-sending"
  | "external-network-call";

export interface SafetyWarning {
  category: WarningCategory;
  message: string;
  matchedText: string;
}

interface Pattern {
  category: WarningCategory;
  regex: RegExp;
  message: string;
}

const PATTERNS: Pattern[] = [
  {
    category: "file-deletion",
    regex: /\bKill\b/g,
    message: "Uses Kill, which permanently deletes a file from disk.",
  },
  {
    category: "file-deletion",
    regex: /\bRmDir\b/g,
    message: "Uses RmDir, which permanently deletes a directory.",
  },
  {
    category: "file-deletion",
    regex: /\.DeleteFile\b/g,
    message: "Uses FileSystemObject.DeleteFile to delete a file from disk.",
  },
  {
    category: "shell-execution",
    regex: /\bShell\s*\(/g,
    message: "Uses Shell to launch an external process or command.",
  },
  {
    category: "shell-execution",
    regex: /\bCreateObject\s*\(\s*"WScript\.Shell"\s*\)/gi,
    message: "Creates a WScript.Shell object, which can run external commands.",
  },
  {
    category: "bulk-row-or-column-deletion",
    regex: /\.EntireRow\.Delete\b/g,
    message: "Deletes entire rows, which can remove data outside the intended range.",
  },
  {
    category: "bulk-row-or-column-deletion",
    regex: /\.EntireColumn\.Delete\b/g,
    message: "Deletes entire columns, which can remove data outside the intended range.",
  },
  {
    category: "bulk-row-or-column-deletion",
    regex: /\bRows\s*\([^)]*\)\.Delete\b/g,
    message: "Bulk-deletes rows.",
  },
  {
    category: "bulk-row-or-column-deletion",
    regex: /\bColumns\s*\([^)]*\)\.Delete\b/g,
    message: "Bulk-deletes columns.",
  },
  {
    category: "workbook-overwrite",
    regex: /\.SaveAs\b/g,
    message: "Uses SaveAs, which can overwrite an existing workbook file.",
  },
  {
    category: "workbook-overwrite",
    regex: /\bApplication\.DisplayAlerts\s*=\s*False\b[\s\S]{0,200}?\.Save\b/g,
    message: "Saves a workbook with alerts suppressed, which can silently overwrite a file.",
  },
  {
    category: "workbook-overwrite",
    regex: /(?<!Save)(?<!SaveAs)\.Save\b(?!\w)/g,
    message: "Saves a workbook, overwriting its file on disk.",
  },
  {
    category: "email-sending",
    regex: /\bCreateObject\s*\(\s*"Outlook\.Application"\s*\)/gi,
    message: "Automates Outlook, which can send email on your behalf.",
  },
  {
    category: "email-sending",
    regex: /\.SendMail\b/g,
    message: "Calls SendMail, which sends an email message.",
  },
  {
    category: "email-sending",
    regex: /\bOutlook\b/g,
    message: "References Outlook automation.",
  },
  {
    category: "external-network-call",
    regex: /\bWinHttp\b/gi,
    message: "Uses WinHTTP to make an external network call.",
  },
  {
    category: "external-network-call",
    regex: /\bXMLHTTP\b/gi,
    message: "Uses XMLHTTP to make an external network call.",
  },
  {
    category: "external-network-call",
    regex: /\bURLDownloadToFile\b/gi,
    message: "Uses URLDownloadToFile to download content from the internet.",
  },
];

/**
 * Scans VBA source text for known destructive/external patterns. Returns
 * one warning per match. This function is pure and side-effect free.
 *
 * IMPORTANT: an empty result does NOT mean the code is safe -- it only
 * means none of the known patterns were found. Always display that caveat
 * next to any use of this function's output.
 */
export function scanVbaForWarnings(vbaCode: string): SafetyWarning[] {
  if (!vbaCode) return [];
  const warnings: SafetyWarning[] = [];

  for (const pattern of PATTERNS) {
    const regex = new RegExp(pattern.regex.source, pattern.regex.flags);
    let match: RegExpExecArray | null;
    while ((match = regex.exec(vbaCode)) !== null) {
      warnings.push({
        category: pattern.category,
        message: pattern.message,
        matchedText: match[0],
      });
      if (match[0].length === 0) {
        regex.lastIndex++;
      }
    }
  }

  return warnings;
}

export const SAFETY_DISCLAIMER =
  "This is an automated pattern scan, not a safety proof. No warning does not mean the code is safe. Always review the full macro yourself and test on a copy of your workbook first.";
