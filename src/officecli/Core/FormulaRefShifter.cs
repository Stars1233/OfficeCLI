// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.RegularExpressions;

namespace OfficeCli.Core;

/// <summary>
/// Direction of insertion that triggered a formula reference shift.
/// </summary>
public enum FormulaShiftDirection
{
    /// <summary>A column was inserted; cell-ref columns at or past insertIdx shift right by 1.</summary>
    ColumnsRight,
    /// <summary>A row was inserted; cell-ref rows at or past insertIdx shift down by 1.</summary>
    RowsDown,
}

/// <summary>
/// Rewrites Excel formula text after a column or row was inserted, so that
/// references that previously pointed to a moved cell continue to point to
/// the same cell.
///
/// <para>This is the regex-based "good enough" implementation (Path A). It
/// handles the common ~90% of formulas: A1 / $A$1 / $A1 / A$1 single refs,
/// A1:B5 ranges, sheet-qualified refs (Sheet2!A1, 'Sheet With Spaces'!A1),
/// and skips string literals and structured-ref bracket content. It does
/// NOT handle: cross-workbook refs ([Book]Sheet!A1), R1C1 notation,
/// whole-column (A:A) or whole-row (1:1) refs, or structured table refs
/// (Table1[Col1]) — those pass through verbatim.</para>
///
/// <para>The public API is intentionally minimal so a future tokenizer-based
/// implementation (Path B) can replace the body of <see cref="Shift"/>
/// without touching call sites or tests.</para>
/// </summary>
public static class FormulaRefShifter
{
    // One regex matches either a single A1 ref or a range, optionally
    // sheet-qualified. Whole-col / whole-row refs are NOT matched here —
    // they require digits in r1, which is mandatory in this pattern.
    //
    // Capture groups:
    //   sheet  — optional sheet name (with surrounding quotes preserved)
    //   c1, r1 — first cell column letters (with optional leading $) and row digits
    //   c2, r2 — range end (or empty for single-cell)
    private static readonly Regex CellRefPattern = new(
        @"(?<![\w.])" +
        @"(?:(?<sheet>'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?" +
        @"(?<c1>\$?[A-Z]{1,3})(?<r1>\$?\d+)" +
        @"(?::(?<c2>\$?[A-Z]{1,3})(?<r2>\$?\d+))?" +
        // (?![\w(]) — also reject when followed by '(' so that function names
        // shaped like `LOG10` / `ATAN2` (col-letters + row-digits) are not
        // misread as cell refs. Cell refs are never followed by '('.
        @"(?![\w(])",
        RegexOptions.Compiled);

    /// <summary>
    /// Returns the formula text rewritten so that any references targeting
    /// <paramref name="modifiedSheet"/> at or past <paramref name="insertIdx"/>
    /// are shifted by 1 in <paramref name="direction"/>. Refs targeting other
    /// sheets, references inside string literals, and references inside
    /// structured-ref brackets are returned untouched.
    /// </summary>
    /// <param name="formula">Formula text without a leading '=' (matching how
    /// the Excel handler stores <c>CellFormula</c> content).</param>
    /// <param name="currentSheet">Sheet that contains the formula. Used to
    /// resolve unqualified refs.</param>
    /// <param name="modifiedSheet">Sheet on which the insert happened. Refs
    /// shift only when their resolved sheet equals this.</param>
    /// <param name="direction">Whether a column or row was inserted.</param>
    /// <param name="insertIdx">1-based column index (for ColumnsRight) or
    /// 1-based row index (for RowsDown) at which the insert happened.</param>
    /// <returns>The rewritten formula text. Returns the input unchanged when
    /// no refs match the shift criteria.</returns>
    public static string Shift(
        string formula,
        string currentSheet,
        string modifiedSheet,
        FormulaShiftDirection direction,
        int insertIdx)
    {
        if (string.IsNullOrEmpty(formula)) return formula;

        var sb = new StringBuilder(formula.Length);
        int i = 0;
        while (i < formula.Length)
        {
            char ch = formula[i];
            if (ch == '"')
            {
                // Copy a string literal verbatim. Excel escapes embedded
                // quotes by doubling them ("" inside "...").
                sb.Append(ch);
                i++;
                while (i < formula.Length)
                {
                    sb.Append(formula[i]);
                    if (formula[i] == '"')
                    {
                        if (i + 1 < formula.Length && formula[i + 1] == '"')
                        {
                            sb.Append(formula[i + 1]);
                            i += 2;
                            continue;
                        }
                        i++;
                        break;
                    }
                    i++;
                }
            }
            else if (ch == '[')
            {
                // Copy bracket content verbatim — covers structured refs
                // (Table1[Col1], Table1[#Headers]) and cross-workbook prefixes
                // ([Book2]Sheet1!A1). Brackets nest in some structured forms.
                int depth = 0;
                while (i < formula.Length)
                {
                    char c = formula[i];
                    sb.Append(c);
                    if (c == '[') depth++;
                    else if (c == ']') { depth--; if (depth == 0) { i++; break; } }
                    i++;
                }
            }
            else
            {
                int start = i;
                while (i < formula.Length && formula[i] != '"' && formula[i] != '[')
                    i++;
                sb.Append(ShiftRefsInChunk(
                    formula.AsSpan(start, i - start).ToString(),
                    currentSheet, modifiedSheet, direction, insertIdx));
            }
        }
        return sb.ToString();
    }

    private static string ShiftRefsInChunk(
        string chunk, string currentSheet, string modifiedSheet,
        FormulaShiftDirection direction, int insertIdx)
    {
        return CellRefPattern.Replace(chunk, m =>
        {
            var sheetGroup = m.Groups["sheet"].Value;
            string targetSheet;
            if (string.IsNullOrEmpty(sheetGroup))
            {
                targetSheet = currentSheet;
            }
            else if (sheetGroup.StartsWith('\'') && sheetGroup.EndsWith('\''))
            {
                targetSheet = sheetGroup[1..^1].Replace("''", "'");
            }
            else
            {
                targetSheet = sheetGroup;
            }

            if (!targetSheet.Equals(modifiedSheet, StringComparison.OrdinalIgnoreCase))
                return m.Value;

            string c1 = m.Groups["c1"].Value;
            string r1 = m.Groups["r1"].Value;
            string c2 = m.Groups["c2"].Value;
            string r2 = m.Groups["r2"].Value;

            string sheetPrefix = string.IsNullOrEmpty(sheetGroup) ? "" : sheetGroup + "!";

            string newC1 = direction == FormulaShiftDirection.ColumnsRight
                ? ShiftColPart(c1, insertIdx) : c1;
            string newR1 = direction == FormulaShiftDirection.RowsDown
                ? ShiftRowPart(r1, insertIdx) : r1;

            if (string.IsNullOrEmpty(c2))
                return $"{sheetPrefix}{newC1}{newR1}";

            string newC2 = direction == FormulaShiftDirection.ColumnsRight
                ? ShiftColPart(c2, insertIdx) : c2;
            string newR2 = direction == FormulaShiftDirection.RowsDown
                ? ShiftRowPart(r2, insertIdx) : r2;
            return $"{sheetPrefix}{newC1}{newR1}:{newC2}{newR2}";
        });
    }

    private static string ShiftColPart(string colPart, int insertColIdx)
    {
        bool isAbs = colPart.StartsWith('$');
        string letters = isAbs ? colPart[1..] : colPart;
        int idx = ColumnLettersToIndex(letters);
        if (idx < insertColIdx) return colPart;
        return (isAbs ? "$" : "") + IndexToColumnLetters(idx + 1);
    }

    private static string ShiftRowPart(string rowPart, int insertRow)
    {
        bool isAbs = rowPart.StartsWith('$');
        int num = int.Parse(isAbs ? rowPart[1..] : rowPart);
        if (num < insertRow) return rowPart;
        return (isAbs ? "$" : "") + (num + 1);
    }

    // Local copies — keep Core/ free of Handlers/ dependencies so the shifter
    // can be used by any handler or tested in isolation.
    private static int ColumnLettersToIndex(string letters)
    {
        int idx = 0;
        foreach (char c in letters)
            idx = idx * 26 + (char.ToUpperInvariant(c) - 'A' + 1);
        return idx;
    }

    private static string IndexToColumnLetters(int idx)
    {
        var sb = new StringBuilder();
        while (idx > 0)
        {
            int rem = (idx - 1) % 26;
            sb.Insert(0, (char)('A' + rem));
            idx = (idx - 1) / 26;
        }
        return sb.ToString();
    }
}
