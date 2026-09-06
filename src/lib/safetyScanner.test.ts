import { describe, expect, it } from "vitest";
import { scanVbaForWarnings } from "./safetyScanner";

describe("scanVbaForWarnings", () => {
  it("returns no warnings for clean code", () => {
    const code = `Option Explicit\nSub DoNothing()\n  Dim x As Long\n  x = 1\nEnd Sub`;
    expect(scanVbaForWarnings(code)).toEqual([]);
  });

  it("flags Kill", () => {
    const warnings = scanVbaForWarnings('Kill "C:\\temp\\file.txt"');
    expect(warnings.some((w) => w.category === "file-deletion")).toBe(true);
  });

  it("flags RmDir", () => {
    const warnings = scanVbaForWarnings('RmDir "C:\\temp"');
    expect(warnings.some((w) => w.category === "file-deletion")).toBe(true);
  });

  it("flags FileSystemObject DeleteFile", () => {
    const warnings = scanVbaForWarnings("fso.DeleteFile path");
    expect(warnings.some((w) => w.category === "file-deletion")).toBe(true);
  });

  it("flags Shell", () => {
    const warnings = scanVbaForWarnings('Shell("cmd.exe")');
    expect(warnings.some((w) => w.category === "shell-execution")).toBe(true);
  });

  it("flags bulk row deletion", () => {
    expect(scanVbaForWarnings("Rows(1).EntireRow.Delete").some((w) => w.category === "bulk-row-or-column-deletion")).toBe(true);
    expect(scanVbaForWarnings("Rows(\"1:5\").Delete").some((w) => w.category === "bulk-row-or-column-deletion")).toBe(true);
  });

  it("flags bulk column deletion", () => {
    expect(scanVbaForWarnings("Columns(1).EntireColumn.Delete").some((w) => w.category === "bulk-row-or-column-deletion")).toBe(true);
    expect(scanVbaForWarnings('Columns("A:B").Delete').some((w) => w.category === "bulk-row-or-column-deletion")).toBe(true);
  });

  it("flags workbook save/overwrite", () => {
    expect(scanVbaForWarnings("wb.SaveAs fileName").some((w) => w.category === "workbook-overwrite")).toBe(true);
    expect(scanVbaForWarnings("wb.Save").some((w) => w.category === "workbook-overwrite")).toBe(true);
  });

  it("flags Outlook automation and SendMail", () => {
    expect(
      scanVbaForWarnings('CreateObject("Outlook.Application")').some((w) => w.category === "email-sending")
    ).toBe(true);
    expect(scanVbaForWarnings("msg.SendMail").some((w) => w.category === "email-sending")).toBe(true);
  });

  it("flags external network calls", () => {
    expect(scanVbaForWarnings('CreateObject("WinHttp.WinHttpRequest.5.1")').some((w) => w.category === "external-network-call")).toBe(true);
    expect(scanVbaForWarnings("MSXML2.XMLHTTP").some((w) => w.category === "external-network-call")).toBe(true);
    expect(scanVbaForWarnings("URLDownloadToFile 0, url, path, 0, 0").some((w) => w.category === "external-network-call")).toBe(true);
  });

  it("does not crash on empty input", () => {
    expect(scanVbaForWarnings("")).toEqual([]);
  });
});
