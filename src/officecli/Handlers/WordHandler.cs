// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using Vml = DocumentFormat.OpenXml.Vml;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public class WordHandler : IDocumentHandler
{
    private readonly WordprocessingDocument _doc;
    private readonly string _filePath;

    public WordHandler(string filePath, bool editable)
    {
        _filePath = filePath;
        _doc = WordprocessingDocument.Open(filePath, editable);
    }

    // ==================== Semantic Layer ====================

    public string ViewAsText(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return "(empty document)";

        var sb = new StringBuilder();
        int lineNum = 0;
        int emitted = 0;
        var bodyElements = GetBodyElements(body).ToList();
        int totalElements = bodyElements.Count;

        foreach (var element in bodyElements)
        {
            lineNum++;
            if (startLine.HasValue && lineNum < startLine.Value) continue;
            if (endLine.HasValue && lineNum > endLine.Value) break;

            if (maxLines.HasValue && emitted >= maxLines.Value)
            {
                sb.AppendLine($"... (showed {emitted} rows, {totalElements} total, use --start/--end to view more)");
                break;
            }

            if (element is Paragraph para)
            {
                // Check if paragraph contains display equation (oMathPara)
                var oMathParaChild = para.ChildElements.FirstOrDefault(e => e.LocalName == "oMathPara" || e is M.Paragraph);
                if (oMathParaChild != null)
                {
                    var mathText = FormulaParser.ToReadableText(oMathParaChild);
                    sb.AppendLine($"[{lineNum}] [Equation] {mathText}");
                }
                else
                {
                    // Check for inline math
                    var mathElements = FindMathElements(para);
                    if (mathElements.Count > 0 && string.IsNullOrWhiteSpace(GetParagraphText(para)))
                    {
                        var mathText = string.Concat(mathElements.Select(FormulaParser.ToReadableText));
                        sb.AppendLine($"[{lineNum}] [Equation] {mathText}");
                    }
                    else if (mathElements.Count > 0)
                    {
                        var text = GetParagraphTextWithMath(para);
                        var listPrefix = GetListPrefix(para);
                        sb.AppendLine($"[{lineNum}] {listPrefix}{text}");
                    }
                    else
                    {
                        var text = GetParagraphText(para);
                        var listPrefix = GetListPrefix(para);
                        sb.AppendLine($"[{lineNum}] {listPrefix}{text}");
                    }
                }
            }
            else if (element.LocalName == "oMathPara" || element is M.Paragraph)
            {
                var mathText = FormulaParser.ToReadableText(element);
                sb.AppendLine($"[{lineNum}] [Equation] {mathText}");
            }
            else if (element is Table table)
            {
                sb.AppendLine($"[{lineNum}] [Table: {table.Elements<TableRow>().Count()} rows]");
            }
            else if (IsStructuralElement(element))
            {
                sb.AppendLine($"[{lineNum}] [{element.LocalName}]");
            }
            else
            {
                // Skip non-content elements (bookmarkStart, bookmarkEnd, proofErr, etc.)
                lineNum--;
                continue;
            }
            emitted++;
        }

        return sb.ToString().TrimEnd();
    }

    public string ViewAsAnnotated(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return "(empty document)";

        var sb = new StringBuilder();
        int lineNum = 0;
        int emitted = 0;
        var bodyElements = GetBodyElements(body).ToList();
        int totalElements = bodyElements.Count;

        foreach (var element in bodyElements)
        {
            lineNum++;
            if (startLine.HasValue && lineNum < startLine.Value) continue;
            if (endLine.HasValue && lineNum > endLine.Value) break;

            if (maxLines.HasValue && emitted >= maxLines.Value)
            {
                sb.AppendLine($"... (showed {emitted} rows, {totalElements} total, use --start/--end to view more)");
                break;
            }

            if (element.LocalName == "oMathPara" || element is M.Paragraph)
            {
                var latex = FormulaParser.ToLatex(element);
                sb.AppendLine($"[{lineNum}] [Equation: \"{latex}\"] ← display");
            }
            else if (element is Paragraph para)
            {
                // Check if paragraph contains display equation (oMathPara)
                var oMathParaChild = para.ChildElements.FirstOrDefault(e => e.LocalName == "oMathPara" || e is M.Paragraph);
                if (oMathParaChild != null)
                {
                    var latex = FormulaParser.ToLatex(oMathParaChild);
                    sb.AppendLine($"[{lineNum}] [Equation: \"{latex}\"] ← display");
                    emitted++;
                    continue;
                }

                var styleName = GetStyleName(para);
                var runs = GetAllRuns(para);

                // Check for inline math
                var inlineMath = FindMathElements(para);
                if (inlineMath.Count > 0 && runs.Count == 0)
                {
                    var latex = string.Concat(inlineMath.Select(FormulaParser.ToLatex));
                    sb.AppendLine($"[{lineNum}] [Equation: \"{latex}\"] ← {styleName} | inline");
                    emitted++;
                    continue;
                }

                if (runs.Count == 0 && inlineMath.Count == 0)
                {
                    sb.AppendLine($"[{lineNum}] [] <- {styleName} | empty paragraph");
                    emitted++;
                    continue;
                }

                var listPrefix = GetListPrefix(para);

                foreach (var run in runs)
                {
                    // Check if run contains an image
                    var drawing = run.GetFirstChild<Drawing>();
                    if (drawing != null)
                    {
                        var imgInfo = GetDrawingInfo(drawing);
                        sb.AppendLine($"[{lineNum}] {listPrefix}[Image: {imgInfo}] ← {styleName}");
                        continue;
                    }

                    var text = GetRunText(run);
                    var fmt = GetRunFormatDescription(run, para);
                    var warn = "";

                    sb.AppendLine($"[{lineNum}] {listPrefix}「{text}」 ← {styleName} | {fmt}{warn}");
                }

                // Show inline math elements
                foreach (var math in inlineMath)
                {
                    var latex = FormulaParser.ToLatex(math);
                    sb.AppendLine($"[{lineNum}] {listPrefix}[Equation: \"{latex}\"] ← {styleName} | inline");
                }
            }
            else if (element is Table table)
            {
                var rows = table.Elements<TableRow>().Count();
                var colCount = table.Elements<TableRow>().FirstOrDefault()
                    ?.Elements<TableCell>().Count() ?? 0;
                sb.AppendLine($"[{lineNum}] [Table: {rows}×{colCount}]");
            }
            emitted++;
        }

        return sb.ToString().TrimEnd();
    }

    public string ViewAsOutline()
    {
        var sb = new StringBuilder();
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return "(empty document)";

        // Document info
        var paragraphs = GetBodyElements(body).OfType<Paragraph>().ToList();
        var tables = GetBodyElements(body).OfType<Table>().ToList();
        var imageCount = body.Descendants<Drawing>().Count();
        var equationCount = body.Descendants().Count(e => e.LocalName == "oMathPara" || e is M.Paragraph);
        var statsLine = $"File: {Path.GetFileName(_filePath)} | {paragraphs.Count} paragraphs | {tables.Count} tables | {imageCount} images";
        if (equationCount > 0) statsLine += $" | {equationCount} equations";
        sb.AppendLine(statsLine);

        // Watermark
        var watermark = FindWatermark();
        if (watermark != null)
            sb.AppendLine($"Watermark: \"{watermark}\"");

        // Headers
        var headers = GetHeaderTexts();
        foreach (var h in headers)
            sb.AppendLine($"Header: \"{h}\"");

        // Footers
        var footers = GetFooterTexts();
        foreach (var f in footers)
            sb.AppendLine($"Footer: \"{f}\"");

        sb.AppendLine();

        // Heading structure
        int lineNum = 0;
        foreach (var para in paragraphs)
        {
            lineNum++;
            var styleName = GetStyleName(para);
            var text = GetParagraphText(para);

            if (styleName.Contains("Heading") || styleName.Contains("标题")
                || styleName.StartsWith("heading", StringComparison.OrdinalIgnoreCase)
                || styleName == "Title" || styleName == "Subtitle")
            {
                var level = GetHeadingLevel(styleName);
                var indent = level <= 1 ? "" : new string(' ', (level - 1) * 2);
                var prefix = level == 0 ? "■" : "├──";
                sb.AppendLine($"{indent}{prefix} [{lineNum}] \"{text}\" ({styleName})");
            }
        }

        return sb.ToString().TrimEnd();
    }

    public string ViewAsStats()
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return "(empty document)";

        var sb = new StringBuilder();
        var paragraphs = GetBodyElements(body).OfType<Paragraph>().ToList();

        // Style counts
        var styleCounts = new Dictionary<string, int>();
        var fontCounts = new Dictionary<string, int>();
        var sizeCounts = new Dictionary<string, int>();
        int emptyParagraphs = 0;
        int doubleSpaces = 0;
        int totalChars = 0;

        foreach (var para in paragraphs)
        {
            var style = GetStyleName(para);
            styleCounts[style] = styleCounts.GetValueOrDefault(style) + 1;

            var runs = GetAllRuns(para);
            if (runs.Count == 0 && string.IsNullOrWhiteSpace(GetParagraphText(para)))
            {
                emptyParagraphs++;
                continue;
            }

            foreach (var run in runs)
            {
                var text = GetRunText(run);
                totalChars += text.Length;

                if (text.Contains("  "))
                    doubleSpaces++;

                var resolved = ResolveEffectiveRunProperties(run, para);
                var font = GetFontFromProperties(resolved) ?? "(default)";
                fontCounts[font] = fontCounts.GetValueOrDefault(font) + 1;

                var size = GetSizeFromProperties(resolved) ?? "(default)";
                sizeCounts[size] = sizeCounts.GetValueOrDefault(size) + 1;
            }
        }

        sb.AppendLine($"Paragraphs: {paragraphs.Count} | Total Characters: {totalChars}");
        sb.AppendLine();

        sb.AppendLine("Style Distribution:");
        foreach (var (style, count) in styleCounts.OrderByDescending(kv => kv.Value))
            sb.AppendLine($"  {style}: {count}");

        sb.AppendLine();
        sb.AppendLine("Font Usage:");
        foreach (var (font, count) in fontCounts.OrderByDescending(kv => kv.Value))
            sb.AppendLine($"  {font}: {count}");

        sb.AppendLine();
        sb.AppendLine("Font Size Usage:");
        foreach (var (size, count) in sizeCounts.OrderByDescending(kv => kv.Value))
            sb.AppendLine($"  {size}: {count}");

        sb.AppendLine();
        sb.AppendLine($"Empty Paragraphs: {emptyParagraphs}");
        sb.AppendLine($"Consecutive Spaces: {doubleSpaces}");

        return sb.ToString().TrimEnd();
    }

    public List<DocumentIssue> ViewAsIssues(string? issueType = null, int? limit = null)
    {
        var issues = new List<DocumentIssue>();
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return issues;

        int issueNum = 0;
        int lineNum = -1;

        foreach (var para in GetBodyElements(body).OfType<Paragraph>())
        {
            lineNum++;
            var styleName = GetStyleName(para);
            var runs = GetAllRuns(para);

            // Empty paragraph
            if (runs.Count == 0 && string.IsNullOrWhiteSpace(GetParagraphText(para)))
            {
                issues.Add(new DocumentIssue
                {
                    Id = $"S{++issueNum}",
                    Type = IssueType.Structure,
                    Severity = IssueSeverity.Warning,
                    Path = $"/body/p[{lineNum + 1}]",
                    Message = "Empty paragraph"
                });
            }

            // Paragraph format checks
            var pProps = para.ParagraphProperties;
            if (pProps != null && IsNormalStyle(styleName))
            {
                var indent = pProps.Indentation;
                if (indent?.FirstLine == null || indent.FirstLine.Value == "0")
                {
                    // Only flag if there's actual text
                    if (runs.Any(r => !string.IsNullOrWhiteSpace(GetRunText(r))))
                    {
                        issues.Add(new DocumentIssue
                        {
                            Id = $"F{++issueNum}",
                            Type = IssueType.Format,
                            Severity = IssueSeverity.Warning,
                            Path = $"/body/p[{lineNum + 1}]",
                            Message = "Body paragraph missing first-line indent",
                            Suggestion = "Set first-line indent to 2 characters"
                        });
                    }
                }
            }

            int runIdx = 0;
            foreach (var run in runs)
            {
                var text = GetRunText(run);

                // Double spaces
                if (text.Contains("  "))
                {
                    issues.Add(new DocumentIssue
                    {
                        Id = $"C{++issueNum}",
                        Type = IssueType.Content,
                        Severity = IssueSeverity.Warning,
                        Path = $"/body/p[{lineNum + 1}]/r[{runIdx + 1}]",
                        Message = "Consecutive spaces",
                        Context = text,
                        Suggestion = "Merge into a single space"
                    });
                }

                // Duplicate punctuation
                if (System.Text.RegularExpressions.Regex.IsMatch(text, @"[，。！？、；：]{2,}"))
                {
                    issues.Add(new DocumentIssue
                    {
                        Id = $"C{++issueNum}",
                        Type = IssueType.Content,
                        Severity = IssueSeverity.Warning,
                        Path = $"/body/p[{lineNum + 1}]/r[{runIdx + 1}]",
                        Message = "Duplicate punctuation",
                        Context = text
                    });
                }

                // Mixed Chinese/English punctuation
                if (HasMixedPunctuation(text))
                {
                    issues.Add(new DocumentIssue
                    {
                        Id = $"C{++issueNum}",
                        Type = IssueType.Content,
                        Severity = IssueSeverity.Info,
                        Path = $"/body/p[{lineNum + 1}]/r[{runIdx + 1}]",
                        Message = "Mixed CJK/Latin punctuation",
                        Context = text
                    });
                }

                runIdx++;
            }

            if (limit.HasValue && issues.Count >= limit.Value) break;
        }

        // Filter by type
        if (issueType != null)
        {
            var type = issueType.ToLowerInvariant() switch
            {
                "format" or "f" => IssueType.Format,
                "content" or "c" => IssueType.Content,
                "structure" or "s" => IssueType.Structure,
                _ => (IssueType?)null
            };
            if (type.HasValue)
                issues = issues.Where(i => i.Type == type.Value).ToList();
        }

        return limit.HasValue ? issues.Take(limit.Value).ToList() : issues;
    }

    // ==================== Query Layer ====================

    public DocumentNode Get(string path, int depth = 1)
    {
        if (path == "/" || path == "")
            return GetRootNode(depth);

        var parts = ParsePath(path);
        var element = NavigateToElement(parts);
        if (element == null)
            return new DocumentNode { Path = path, Type = "error", Text = $"Path not found: {path}" };

        return ElementToNode(element, path, depth);
    }

    public List<DocumentNode> Query(string selector)
    {
        var results = new List<DocumentNode>();
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return results;

        // Simple selector parser: element[attr=value]
        var parsed = ParseSelector(selector);

        // Determine if main selector targets runs directly (no > parent)
        bool isRunSelector = parsed.ChildSelector == null &&
            (parsed.Element == "r" || parsed.Element == "run");
        bool isPictureSelector = parsed.ChildSelector == null &&
            (parsed.Element == "picture" || parsed.Element == "image" || parsed.Element == "img");
        bool isEquationSelector = parsed.ChildSelector == null &&
            (parsed.Element == "equation" || parsed.Element == "math" || parsed.Element == "formula");

        // Scheme B: generic XML fallback for unrecognized element types
        // Use GenericXmlQuery.ParseSelector which properly handles namespace prefixes (e.g., "a:ln")
        var genericParsed = GenericXmlQuery.ParseSelector(selector);
        bool isKnownType = string.IsNullOrEmpty(genericParsed.element)
            || genericParsed.element is "p" or "paragraph" or "r" or "run"
                or "picture" or "image" or "img"
                or "equation" or "math" or "formula";
        if (!isKnownType && parsed.ChildSelector == null)
        {
            var root = _doc.MainDocumentPart?.Document;
            if (root != null)
                return GenericXmlQuery.Query(root, genericParsed.element, genericParsed.attrs, genericParsed.containsText);
            return results;
        }

        int paraIdx = -1;
        int mathParaIdx = -1;
        foreach (var element in body.ChildElements)
        {
            // Display equations (m:oMathPara) at body level
            if (element.LocalName == "oMathPara" || element is M.Paragraph)
            {
                mathParaIdx++;
                if (isEquationSelector)
                {
                    var latex = FormulaParser.ToLatex(element);
                    if (parsed.ContainsText == null || latex.Contains(parsed.ContainsText))
                    {
                        results.Add(new DocumentNode
                        {
                            Path = $"/body/oMathPara[{mathParaIdx + 1}]",
                            Type = "equation",
                            Text = latex,
                            Format = { ["mode"] = "display" }
                        });
                    }
                }
                continue;
            }

            if (element is Paragraph para)
            {
                paraIdx++;

                if (isEquationSelector)
                {
                    // Check for display equation (oMathPara inside w:p)
                    var oMathParaInPara = para.ChildElements.FirstOrDefault(e => e.LocalName == "oMathPara" || e is M.Paragraph);
                    if (oMathParaInPara != null)
                    {
                        mathParaIdx++;
                        var latex = FormulaParser.ToLatex(oMathParaInPara);
                        if (parsed.ContainsText == null || latex.Contains(parsed.ContainsText))
                        {
                            results.Add(new DocumentNode
                            {
                                Path = $"/body/oMathPara[{mathParaIdx + 1}]",
                                Type = "equation",
                                Text = latex,
                                Format = { ["mode"] = "display" }
                            });
                        }
                        continue;
                    }

                    // Find inline math in this paragraph
                    int mathIdx = 0;
                    foreach (var oMath in para.ChildElements.Where(e => e.LocalName == "oMath" || e is M.OfficeMath))
                    {
                        var latex = FormulaParser.ToLatex(oMath);
                        if (parsed.ContainsText == null || latex.Contains(parsed.ContainsText))
                        {
                            results.Add(new DocumentNode
                            {
                                Path = $"/body/p[{paraIdx + 1}]/oMath[{mathIdx + 1}]",
                                Type = "equation",
                                Text = latex,
                                Format = { ["mode"] = "inline" }
                            });
                        }
                        mathIdx++;
                    }
                }
                else if (isPictureSelector)
                {
                    int runIdx = 0;
                    foreach (var run in GetAllRuns(para))
                    {
                        var drawing = run.GetFirstChild<Drawing>();
                        if (drawing != null)
                        {
                            bool noAlt = parsed.Attributes.ContainsKey("__no-alt");
                            if (noAlt)
                            {
                                var docProps = drawing.Descendants<DW.DocProperties>().FirstOrDefault();
                                if (string.IsNullOrEmpty(docProps?.Description?.Value))
                                    results.Add(CreateImageNode(drawing, run, $"/body/p[{paraIdx + 1}]/r[{runIdx + 1}]"));
                            }
                            else
                            {
                                results.Add(CreateImageNode(drawing, run, $"/body/p[{paraIdx + 1}]/r[{runIdx + 1}]"));
                            }
                        }
                        runIdx++;
                    }
                }
                else if (isRunSelector)
                {
                    // Main selector targets runs: search all runs in all paragraphs
                    int runIdx = 0;
                    foreach (var run in GetAllRuns(para))
                    {
                        if (MatchesRunSelector(run, para, parsed))
                        {
                            results.Add(ElementToNode(run, $"/body/p[{paraIdx + 1}]/r[{runIdx + 1}]", 0));
                        }
                        runIdx++;
                    }
                }
                else
                {
                    if (MatchesSelector(para, parsed, paraIdx))
                    {
                        results.Add(ElementToNode(para, $"/body/p[{paraIdx + 1}]", 0));
                    }

                    if (parsed.ChildSelector != null)
                    {
                        int runIdx = 0;
                        foreach (var run in GetAllRuns(para))
                        {
                            if (MatchesRunSelector(run, para, parsed.ChildSelector))
                            {
                                results.Add(ElementToNode(run, $"/body/p[{paraIdx + 1}]/r[{runIdx + 1}]", 0));
                            }
                            runIdx++;
                        }
                    }
                }
            }
        }

        return results;
    }

    public List<string> Set(string path, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();

        // Document-level properties
        if (path == "/" || path == "")
        {
            SetDocumentProperties(properties);
            _doc.MainDocumentPart?.Document?.Save();
            return unsupported;
        }

        var parts = ParsePath(path);
        var element = NavigateToElement(parts);
        if (element == null)
            throw new ArgumentException($"Path not found: {path}");

        if (element is Run run)
        {
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "text":
                        var textEl = run.GetFirstChild<Text>();
                        if (textEl != null) textEl.Text = value;
                        break;
                    case "bold":
                        EnsureRunProperties(run).Bold = bool.Parse(value) ? new Bold() : null;
                        break;
                    case "italic":
                        EnsureRunProperties(run).Italic = bool.Parse(value) ? new Italic() : null;
                        break;
                    case "caps":
                        EnsureRunProperties(run).Caps = bool.Parse(value) ? new Caps() : null;
                        break;
                    case "smallcaps":
                        EnsureRunProperties(run).SmallCaps = bool.Parse(value) ? new SmallCaps() : null;
                        break;
                    case "dstrike":
                        EnsureRunProperties(run).DoubleStrike = bool.Parse(value) ? new DoubleStrike() : null;
                        break;
                    case "vanish":
                        EnsureRunProperties(run).Vanish = bool.Parse(value) ? new Vanish() : null;
                        break;
                    case "outline":
                        EnsureRunProperties(run).Outline = bool.Parse(value) ? new Outline() : null;
                        break;
                    case "shadow":
                        EnsureRunProperties(run).Shadow = bool.Parse(value) ? new Shadow() : null;
                        break;
                    case "emboss":
                        EnsureRunProperties(run).Emboss = bool.Parse(value) ? new Emboss() : null;
                        break;
                    case "imprint":
                        EnsureRunProperties(run).Imprint = bool.Parse(value) ? new Imprint() : null;
                        break;
                    case "noproof":
                        EnsureRunProperties(run).NoProof = bool.Parse(value) ? new NoProof() : null;
                        break;
                    case "rtl":
                        EnsureRunProperties(run).RightToLeftText = bool.Parse(value) ? new RightToLeftText() : null;
                        break;
                    case "font":
                        var rPrFont = EnsureRunProperties(run);
                        var existingFonts = rPrFont.RunFonts;
                        if (existingFonts != null)
                        {
                            existingFonts.Ascii = value;
                            existingFonts.HighAnsi = value;
                            existingFonts.EastAsia = value;
                        }
                        else
                        {
                            rPrFont.RunFonts = new RunFonts { Ascii = value, HighAnsi = value, EastAsia = value };
                        }
                        break;
                    case "size":
                        EnsureRunProperties(run).FontSize = new FontSize
                        {
                            Val = (int.Parse(value) * 2).ToString() // half-points
                        };
                        break;
                    case "highlight":
                        EnsureRunProperties(run).Highlight = new Highlight
                        {
                            Val = new HighlightColorValues(value)
                        };
                        break;
                    case "color":
                        EnsureRunProperties(run).Color = new Color { Val = value };
                        break;
                    case "underline":
                        EnsureRunProperties(run).Underline = new Underline
                        {
                            Val = new UnderlineValues(value)
                        };
                        break;
                    case "strike":
                        EnsureRunProperties(run).Strike = bool.Parse(value) ? new Strike() : null;
                        break;
                    case "shading":
                    case "shd":
                        // shd has w:val, w:fill, w:color — value format: "fill" or "val;fill" or "val;fill;color"
                        var shdParts = value.Split(';');
                        var shd = new Shading();
                        if (shdParts.Length == 1)
                        {
                            shd.Val = ShadingPatternValues.Clear;
                            shd.Fill = shdParts[0];
                        }
                        else if (shdParts.Length >= 2)
                        {
                            shd.Val = new ShadingPatternValues(shdParts[0]);
                            shd.Fill = shdParts[1];
                            if (shdParts.Length >= 3) shd.Color = shdParts[2];
                        }
                        EnsureRunProperties(run).Shading = shd;
                        break;
                    case "alt":
                        var drawingAlt = run.GetFirstChild<Drawing>();
                        if (drawingAlt != null)
                        {
                            var docPropsAlt = drawingAlt.Descendants<DW.DocProperties>().FirstOrDefault();
                            if (docPropsAlt != null) docPropsAlt.Description = value;
                        }
                        else unsupported.Add(key);
                        break;
                    case "width":
                        var drawingW = run.GetFirstChild<Drawing>();
                        if (drawingW != null)
                        {
                            var extentW = drawingW.Descendants<DW.Extent>().FirstOrDefault();
                            if (extentW != null) extentW.Cx = ParseEmu(value);
                            var extentsW = drawingW.Descendants<A.Extents>().FirstOrDefault();
                            if (extentsW != null) extentsW.Cx = ParseEmu(value);
                        }
                        else unsupported.Add(key);
                        break;
                    case "height":
                        var drawingH = run.GetFirstChild<Drawing>();
                        if (drawingH != null)
                        {
                            var extentH = drawingH.Descendants<DW.Extent>().FirstOrDefault();
                            if (extentH != null) extentH.Cy = ParseEmu(value);
                            var extentsH = drawingH.Descendants<A.Extents>().FirstOrDefault();
                            if (extentsH != null) extentsH.Cy = ParseEmu(value);
                        }
                        else unsupported.Add(key);
                        break;
                    default:
                        if (!GenericXmlQuery.TryCreateTypedChild(EnsureRunProperties(run), key, value))
                            unsupported.Add(key);
                        break;
                }
            }
        }
        else if (element is Paragraph para)
        {
            var pProps = para.ParagraphProperties ?? para.PrependChild(new ParagraphProperties());
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "style":
                        pProps.ParagraphStyleId = new ParagraphStyleId { Val = value };
                        break;
                    case "alignment":
                        pProps.Justification = new Justification
                        {
                            Val = value.ToLowerInvariant() switch
                            {
                                "center" => JustificationValues.Center,
                                "right" => JustificationValues.Right,
                                "justify" => JustificationValues.Both,
                                _ => JustificationValues.Left
                            }
                        };
                        break;
                    case "firstlineindent":
                        var indent = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                        indent.FirstLine = (int.Parse(value) * 480).ToString(); // chars to twips (~480 per char)
                        break;
                    case "shading":
                    case "shd":
                        var shdPartsP = value.Split(';');
                        var shdP = new Shading();
                        if (shdPartsP.Length == 1)
                        {
                            shdP.Val = ShadingPatternValues.Clear;
                            shdP.Fill = shdPartsP[0];
                        }
                        else if (shdPartsP.Length >= 2)
                        {
                            shdP.Val = new ShadingPatternValues(shdPartsP[0]);
                            shdP.Fill = shdPartsP[1];
                            if (shdPartsP.Length >= 3) shdP.Color = shdPartsP[2];
                        }
                        pProps.Shading = shdP;
                        break;
                    case "spacebefore":
                        var spacingBefore = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                        spacingBefore.Before = value;
                        break;
                    case "spaceafter":
                        var spacingAfter = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                        spacingAfter.After = value;
                        break;
                    case "linespacing":
                        var spacingLine = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                        spacingLine.Line = value;
                        spacingLine.LineRule = LineSpacingRuleValues.Auto;
                        break;
                    case "numid":
                        var numPr = pProps.NumberingProperties ?? (pProps.NumberingProperties = new NumberingProperties());
                        numPr.NumberingId = new NumberingId { Val = int.Parse(value) };
                        break;
                    case "numlevel" or "ilvl":
                        var numPr2 = pProps.NumberingProperties ?? (pProps.NumberingProperties = new NumberingProperties());
                        numPr2.NumberingLevelReference = new NumberingLevelReference { Val = int.Parse(value) };
                        break;
                    case "liststyle":
                        ApplyListStyle(para, value);
                        break;
                    default:
                        if (!GenericXmlQuery.TryCreateTypedChild(pProps, key, value))
                            unsupported.Add(key);
                        break;
                }
            }
        }

        else if (element is TableCell cell)
        {
            var tcPr = cell.TableCellProperties ?? cell.PrependChild(new TableCellProperties());
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "text":
                        var firstPara = cell.Elements<Paragraph>().FirstOrDefault();
                        if (firstPara == null)
                        {
                            firstPara = new Paragraph();
                            cell.AppendChild(firstPara);
                        }
                        // Remove existing runs
                        foreach (var r in firstPara.Elements<Run>().ToList()) r.Remove();
                        firstPara.AppendChild(new Run(new Text(value) { Space = SpaceProcessingModeValues.Preserve }));
                        break;
                    case "font":
                    case "size":
                    case "bold":
                    case "italic":
                    case "color":
                        // Apply to all runs in all paragraphs in the cell
                        foreach (var cellPara in cell.Elements<Paragraph>())
                        {
                            foreach (var cellRun in cellPara.Elements<Run>())
                            {
                                var rPr = EnsureRunProperties(cellRun);
                                switch (key.ToLowerInvariant())
                                {
                                    case "font":
                                        rPr.RunFonts = new RunFonts { Ascii = value, HighAnsi = value, EastAsia = value };
                                        break;
                                    case "size":
                                        rPr.FontSize = new FontSize { Val = (int.Parse(value) * 2).ToString() };
                                        break;
                                    case "bold":
                                        rPr.Bold = bool.Parse(value) ? new Bold() : null;
                                        break;
                                    case "italic":
                                        rPr.Italic = bool.Parse(value) ? new Italic() : null;
                                        break;
                                    case "color":
                                        rPr.Color = new Color { Val = value };
                                        break;
                                }
                            }
                        }
                        break;
                    case "shd" or "shading":
                        var shdParts = value.Split(';');
                        var shd = new Shading();
                        if (shdParts.Length == 1)
                        {
                            shd.Val = ShadingPatternValues.Clear;
                            shd.Fill = shdParts[0];
                        }
                        else if (shdParts.Length >= 2)
                        {
                            shd.Val = new ShadingPatternValues(shdParts[0]);
                            shd.Fill = shdParts[1];
                            if (shdParts.Length >= 3) shd.Color = shdParts[2];
                        }
                        tcPr.Shading = shd;
                        break;
                    case "alignment":
                        var cellFirstPara = cell.Elements<Paragraph>().FirstOrDefault();
                        if (cellFirstPara != null)
                        {
                            var cpProps = cellFirstPara.ParagraphProperties ?? cellFirstPara.PrependChild(new ParagraphProperties());
                            cpProps.Justification = new Justification
                            {
                                Val = value.ToLowerInvariant() switch
                                {
                                    "center" => JustificationValues.Center,
                                    "right" => JustificationValues.Right,
                                    "justify" => JustificationValues.Both,
                                    _ => JustificationValues.Left
                                }
                            };
                        }
                        break;
                    case "valign":
                        tcPr.TableCellVerticalAlignment = new TableCellVerticalAlignment
                        {
                            Val = value.ToLowerInvariant() switch
                            {
                                "center" => TableVerticalAlignmentValues.Center,
                                "bottom" => TableVerticalAlignmentValues.Bottom,
                                _ => TableVerticalAlignmentValues.Top
                            }
                        };
                        break;
                    case "width":
                        tcPr.TableCellWidth = new TableCellWidth { Width = value, Type = TableWidthUnitValues.Dxa };
                        break;
                    case "vmerge":
                        tcPr.VerticalMerge = new VerticalMerge
                        {
                            Val = value.ToLowerInvariant() == "restart" ? MergedCellValues.Restart : MergedCellValues.Continue
                        };
                        break;
                    case "gridspan":
                        var newSpan = int.Parse(value);
                        tcPr.GridSpan = new GridSpan { Val = newSpan };
                        // Ensure the row has the correct number of tc elements.
                        // Calculate total grid columns occupied by all cells in this row,
                        // then remove/add cells so it matches the table grid.
                        if (element.Parent is TableRow parentRow)
                        {
                            var table = parentRow.Parent as Table;
                            var gridCols = table?.GetFirstChild<TableGrid>()
                                ?.Elements<GridColumn>().Count() ?? 0;
                            if (gridCols > 0)
                            {
                                // Calculate total columns occupied by current cells
                                var totalSpan = parentRow.Elements<TableCell>().Sum(tc =>
                                    tc.TableCellProperties?.GridSpan?.Val?.Value ?? 1);
                                // Remove excess cells after the current cell
                                while (totalSpan > gridCols)
                                {
                                    var nextCell = ((TableCell)element).NextSibling<TableCell>();
                                    if (nextCell == null) break;
                                    totalSpan -= nextCell.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
                                    nextCell.Remove();
                                }
                            }
                        }
                        break;
                    default:
                        if (!GenericXmlQuery.TryCreateTypedChild(tcPr, key, value))
                            unsupported.Add(key);
                        break;
                }
            }
        }
        else if (element is TableRow row)
        {
            var trPr = row.TableRowProperties ?? row.PrependChild(new TableRowProperties());
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "height":
                        trPr.AppendChild(new TableRowHeight { Val = uint.Parse(value), HeightType = HeightRuleValues.AtLeast });
                        break;
                    case "header":
                        if (bool.Parse(value))
                            trPr.AppendChild(new TableHeader());
                        else
                            trPr.GetFirstChild<TableHeader>()?.Remove();
                        break;
                    default:
                        if (!GenericXmlQuery.TryCreateTypedChild(trPr, key, value))
                            unsupported.Add(key);
                        break;
                }
            }
        }
        else if (element is Table tbl)
        {
            var tblPr = tbl.GetFirstChild<TableProperties>() ?? tbl.PrependChild(new TableProperties());
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "alignment":
                        tblPr.TableJustification = new TableJustification
                        {
                            Val = value.ToLowerInvariant() switch
                            {
                                "center" => TableRowAlignmentValues.Center,
                                "right" => TableRowAlignmentValues.Right,
                                _ => TableRowAlignmentValues.Left
                            }
                        };
                        break;
                    case "width":
                        tblPr.TableWidth = new TableWidth { Width = value, Type = TableWidthUnitValues.Dxa };
                        break;
                    default:
                        if (!GenericXmlQuery.TryCreateTypedChild(tblPr, key, value))
                            unsupported.Add(key);
                        break;
                }
            }
        }

        _doc.MainDocumentPart?.Document?.Save();
        return unsupported;
    }

    public string Add(string parentPath, string type, int? index, Dictionary<string, string> properties)
    {
        var body = _doc.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("Document body not found");

        OpenXmlElement parent;
        if (parentPath is "/" or "" or "/body")
        {
            parent = body;
        }
        else
        {
            var parts = ParsePath(parentPath);
            parent = NavigateToElement(parts)
                ?? throw new ArgumentException($"Path not found: {parentPath}");
        }

        OpenXmlElement newElement;
        string resultPath;

        switch (type.ToLowerInvariant())
        {
            case "paragraph" or "p":
                var para = new Paragraph();
                var pProps = new ParagraphProperties();

                if (properties.TryGetValue("style", out var style))
                    pProps.ParagraphStyleId = new ParagraphStyleId { Val = style };
                if (properties.TryGetValue("alignment", out var alignment))
                    pProps.Justification = new Justification
                    {
                        Val = alignment.ToLowerInvariant() switch
                        {
                            "center" => JustificationValues.Center,
                            "right" => JustificationValues.Right,
                            "justify" => JustificationValues.Both,
                            _ => JustificationValues.Left
                        }
                    };
                if (properties.TryGetValue("firstlineindent", out var indent))
                {
                    pProps.Indentation = new Indentation
                    {
                        FirstLine = (int.Parse(indent) * 480).ToString()
                    };
                }
                if (properties.TryGetValue("spacebefore", out var sb4))
                {
                    var spacing = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                    spacing.Before = sb4;
                }
                if (properties.TryGetValue("spaceafter", out var sa4))
                {
                    var spacing = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                    spacing.After = sa4;
                }
                if (properties.TryGetValue("linespacing", out var ls4))
                {
                    var spacing = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                    spacing.Line = ls4;
                    spacing.LineRule = LineSpacingRuleValues.Auto;
                }
                if (properties.TryGetValue("numid", out var numId))
                {
                    var numPr = pProps.NumberingProperties ?? (pProps.NumberingProperties = new NumberingProperties());
                    numPr.NumberingId = new NumberingId { Val = int.Parse(numId) };
                    if (properties.TryGetValue("numlevel", out var numLevel))
                        numPr.NumberingLevelReference = new NumberingLevelReference { Val = int.Parse(numLevel) };
                }
                if (properties.TryGetValue("shd", out var pShdVal) || properties.TryGetValue("shading", out pShdVal))
                {
                    var shdParts = pShdVal.Split(';');
                    var shd = new Shading();
                    if (shdParts.Length == 1)
                    {
                        shd.Val = ShadingPatternValues.Clear;
                        shd.Fill = shdParts[0];
                    }
                    else if (shdParts.Length >= 2)
                    {
                        shd.Val = new ShadingPatternValues(shdParts[0]);
                        shd.Fill = shdParts[1];
                        if (shdParts.Length >= 3) shd.Color = shdParts[2];
                    }
                    pProps.Shading = shd;
                }
                if (properties.TryGetValue("liststyle", out var listStyle))
                {
                    para.AppendChild(pProps);
                    ApplyListStyle(para, listStyle);
                    // pProps already appended, skip the append below
                    goto paragraphPropsApplied;
                }

                para.AppendChild(pProps);
                paragraphPropsApplied:

                if (properties.TryGetValue("text", out var text))
                {
                    var run = new Run();
                    var rProps = new RunProperties();
                    if (properties.TryGetValue("font", out var font))
                    {
                        rProps.AppendChild(new RunFonts { Ascii = font, HighAnsi = font, EastAsia = font });
                    }
                    if (properties.TryGetValue("size", out var size))
                    {
                        rProps.AppendChild(new FontSize { Val = (int.Parse(size) * 2).ToString() });
                    }
                    if (properties.TryGetValue("bold", out var bold) && bool.Parse(bold))
                        rProps.Bold = new Bold();
                    if (properties.TryGetValue("italic", out var pItalic) && bool.Parse(pItalic))
                        rProps.Italic = new Italic();
                    if (properties.TryGetValue("color", out var pColor))
                        rProps.Color = new Color { Val = pColor };
                    if (properties.TryGetValue("underline", out var pUnderline))
                        rProps.Underline = new Underline { Val = new UnderlineValues(pUnderline) };
                    if (properties.TryGetValue("strike", out var pStrike) && bool.Parse(pStrike))
                        rProps.Strike = new Strike();
                    if (properties.TryGetValue("highlight", out var pHighlight))
                        rProps.Highlight = new Highlight { Val = new HighlightColorValues(pHighlight) };
                    if (properties.TryGetValue("caps", out var pCaps) && bool.Parse(pCaps))
                        rProps.Caps = new Caps();
                    if (properties.TryGetValue("smallcaps", out var pSmallCaps) && bool.Parse(pSmallCaps))
                        rProps.SmallCaps = new SmallCaps();
                    if (properties.TryGetValue("shd", out var pShd) || properties.TryGetValue("shading", out pShd))
                    {
                        var shdParts = pShd.Split(';');
                        var shd = new Shading();
                        if (shdParts.Length == 1)
                        {
                            shd.Val = ShadingPatternValues.Clear;
                            shd.Fill = shdParts[0];
                        }
                        else if (shdParts.Length >= 2)
                        {
                            shd.Val = new ShadingPatternValues(shdParts[0]);
                            shd.Fill = shdParts[1];
                            if (shdParts.Length >= 3) shd.Color = shdParts[2];
                        }
                        rProps.Shading = shd;
                    }

                    run.AppendChild(rProps);
                    run.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
                    para.AppendChild(run);
                }

                newElement = para;
                var paraCount = parent.Elements<Paragraph>().Count();
                if (index.HasValue && index.Value < paraCount)
                {
                    var refElement = parent.Elements<Paragraph>().ElementAt(index.Value);
                    parent.InsertBefore(para, refElement);
                    resultPath = $"{parentPath}/p[{index.Value + 1}]";
                }
                else
                {
                    parent.AppendChild(para);
                    resultPath = $"{parentPath}/p[{paraCount + 1}]";
                }
                break;

            case "equation" or "formula" or "math":
                if (!properties.TryGetValue("formula", out var formula))
                    throw new ArgumentException("'formula' property is required for equation type");

                var mode = properties.GetValueOrDefault("mode", "display");

                if (mode == "inline" && parent is Paragraph inlinePara)
                {
                    // Insert inline math into existing paragraph
                    var mathElement = FormulaParser.Parse(formula);
                    if (mathElement is M.OfficeMath oMathInline)
                        inlinePara.AppendChild(oMathInline);
                    else
                        inlinePara.AppendChild(new M.OfficeMath(mathElement.CloneNode(true)));
                    var mathCount = inlinePara.Elements<M.OfficeMath>().Count();
                    resultPath = $"{parentPath}/oMath[{mathCount}]";
                    newElement = inlinePara;
                }
                else
                {
                    // Display mode: create m:oMathPara
                    var mathContent = FormulaParser.Parse(formula);
                    M.OfficeMath oMath;
                    if (mathContent is M.OfficeMath directMath)
                        oMath = directMath;
                    else
                        oMath = new M.OfficeMath(mathContent.CloneNode(true));

                    var mathPara = new M.Paragraph(oMath);

                    if (parent is Body || parent is SdtBlock)
                    {
                        // Wrap m:oMathPara in w:p for schema validity
                        var wrapPara = new Paragraph(mathPara);
                        var mathParaCount = parent.Descendants<M.Paragraph>().Count();
                        if (index.HasValue)
                        {
                            var children = parent.ChildElements.ToList();
                            if (index.Value < children.Count)
                                parent.InsertBefore(wrapPara, children[index.Value]);
                            else
                                parent.AppendChild(wrapPara);
                        }
                        else
                        {
                            parent.AppendChild(wrapPara);
                        }
                        resultPath = $"{parentPath}/oMathPara[{mathParaCount + 1}]";
                    }
                    else
                    {
                        parent.AppendChild(mathPara);
                        resultPath = $"{parentPath}/oMathPara[1]";
                    }
                    newElement = mathPara;
                }

                _doc.MainDocumentPart?.Document?.Save();
                return resultPath;

            case "run" or "r":
                if (parent is not Paragraph targetPara)
                    throw new ArgumentException("Runs can only be added to paragraphs");

                var newRun = new Run();
                var newRProps = new RunProperties();
                if (properties.TryGetValue("font", out var rFont))
                    newRProps.AppendChild(new RunFonts { Ascii = rFont, HighAnsi = rFont, EastAsia = rFont });
                if (properties.TryGetValue("size", out var rSize))
                    newRProps.AppendChild(new FontSize { Val = (int.Parse(rSize) * 2).ToString() });
                if (properties.TryGetValue("bold", out var rBold) && bool.Parse(rBold))
                    newRProps.Bold = new Bold();
                if (properties.TryGetValue("italic", out var rItalic) && bool.Parse(rItalic))
                    newRProps.Italic = new Italic();
                if (properties.TryGetValue("color", out var rColor))
                    newRProps.Color = new Color { Val = rColor };
                if (properties.TryGetValue("underline", out var rUnderline))
                    newRProps.Underline = new Underline { Val = new UnderlineValues(rUnderline) };
                if (properties.TryGetValue("strike", out var rStrike) && bool.Parse(rStrike))
                    newRProps.Strike = new Strike();
                if (properties.TryGetValue("highlight", out var rHighlight))
                    newRProps.Highlight = new Highlight { Val = new HighlightColorValues(rHighlight) };
                if (properties.TryGetValue("caps", out var rCaps) && bool.Parse(rCaps))
                    newRProps.Caps = new Caps();
                if (properties.TryGetValue("smallcaps", out var rSmallCaps) && bool.Parse(rSmallCaps))
                    newRProps.SmallCaps = new SmallCaps();
                if (properties.TryGetValue("shd", out var rShd) || properties.TryGetValue("shading", out rShd))
                {
                    var shdParts = rShd.Split(';');
                    var shd = new Shading();
                    if (shdParts.Length == 1)
                    {
                        shd.Val = ShadingPatternValues.Clear;
                        shd.Fill = shdParts[0];
                    }
                    else if (shdParts.Length >= 2)
                    {
                        shd.Val = new ShadingPatternValues(shdParts[0]);
                        shd.Fill = shdParts[1];
                        if (shdParts.Length >= 3) shd.Color = shdParts[2];
                    }
                    newRProps.Shading = shd;
                }

                newRun.AppendChild(newRProps);
                var runText = properties.GetValueOrDefault("text", "");
                newRun.AppendChild(new Text(runText) { Space = SpaceProcessingModeValues.Preserve });

                var runCount = targetPara.Elements<Run>().Count();
                if (index.HasValue && index.Value < runCount)
                {
                    var refRun = targetPara.Elements<Run>().ElementAt(index.Value);
                    targetPara.InsertBefore(newRun, refRun);
                    resultPath = $"{parentPath}/r[{index.Value + 1}]";
                }
                else
                {
                    targetPara.AppendChild(newRun);
                    resultPath = $"{parentPath}/r[{runCount + 1}]";
                }

                newElement = newRun;
                break;

            case "table" or "tbl":
                var table = new Table();
                var tblProps = new TableProperties(
                    new TableBorders(
                        new TopBorder { Val = BorderValues.Single, Size = 4 },
                        new BottomBorder { Val = BorderValues.Single, Size = 4 },
                        new StartBorder { Val = BorderValues.Single, Size = 4 },
                        new EndBorder { Val = BorderValues.Single, Size = 4 },
                        new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                        new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                    )
                );
                table.AppendChild(tblProps);

                int rows = properties.TryGetValue("rows", out var rowsStr) ? int.Parse(rowsStr) : 1;
                int cols = properties.TryGetValue("cols", out var colsStr) ? int.Parse(colsStr) : 1;

                // Add table grid
                var tblGrid = new TableGrid();
                for (int gc = 0; gc < cols; gc++)
                    tblGrid.AppendChild(new GridColumn { Width = "2400" });
                table.AppendChild(tblGrid);

                for (int r = 0; r < rows; r++)
                {
                    var row = new TableRow();
                    for (int c = 0; c < cols; c++)
                    {
                        var cell = new TableCell(new Paragraph());
                        row.AppendChild(cell);
                    }
                    table.AppendChild(row);
                }

                parent.AppendChild(table);
                var tblCount = parent.Elements<Table>().Count();
                resultPath = $"{parentPath}/tbl[{tblCount}]";
                newElement = table;
                break;

            case "picture" or "image" or "img":
                if (!properties.TryGetValue("path", out var imgPath))
                    throw new ArgumentException("'path' property is required for picture type");
                if (!File.Exists(imgPath))
                    throw new FileNotFoundException($"Image file not found: {imgPath}");

                var imgExtension = Path.GetExtension(imgPath).ToLowerInvariant();
                var imgPartType = imgExtension switch
                {
                    ".png" => ImagePartType.Png,
                    ".jpg" or ".jpeg" => ImagePartType.Jpeg,
                    ".gif" => ImagePartType.Gif,
                    ".bmp" => ImagePartType.Bmp,
                    ".tif" or ".tiff" => ImagePartType.Tiff,
                    ".emf" => ImagePartType.Emf,
                    ".wmf" => ImagePartType.Wmf,
                    ".svg" => ImagePartType.Svg,
                    _ => throw new ArgumentException($"Unsupported image format: {imgExtension}")
                };

                var mainPart = _doc.MainDocumentPart!;
                var imagePart = mainPart.AddImagePart(imgPartType);
                using (var stream = File.OpenRead(imgPath))
                    imagePart.FeedData(stream);
                var relId = mainPart.GetIdOfPart(imagePart);

                // Determine dimensions (default: 6 inches wide, auto height)
                long cxEmu = 5486400; // 6 inches in EMUs (914400 * 6)
                long cyEmu = 3657600; // 4 inches default
                if (properties.TryGetValue("width", out var widthStr))
                    cxEmu = ParseEmu(widthStr);
                if (properties.TryGetValue("height", out var heightStr))
                    cyEmu = ParseEmu(heightStr);

                var altText = properties.GetValueOrDefault("alt", Path.GetFileName(imgPath));

                var imgRun = CreateImageRun(relId, cxEmu, cyEmu, altText);

                Paragraph imgPara;
                if (parent is Paragraph existingPara)
                {
                    existingPara.AppendChild(imgRun);
                    imgPara = existingPara;
                    var imgRunCount = existingPara.Elements<Run>().Count();
                    resultPath = $"{parentPath}/r[{imgRunCount}]";
                }
                else
                {
                    imgPara = new Paragraph(imgRun);
                    var imgParaCount = parent.Elements<Paragraph>().Count();
                    if (index.HasValue && index.Value < imgParaCount)
                    {
                        var refPara = parent.Elements<Paragraph>().ElementAt(index.Value);
                        parent.InsertBefore(imgPara, refPara);
                        resultPath = $"{parentPath}/p[{index.Value + 1}]";
                    }
                    else
                    {
                        parent.AppendChild(imgPara);
                        resultPath = $"{parentPath}/p[{imgParaCount + 1}]";
                    }
                }
                newElement = imgPara;
                break;

            case "comment":
            {
                if (!properties.TryGetValue("text", out var commentText))
                    throw new ArgumentException("'text' property is required for comment type");

                var commentRun = parent as Run;
                var commentPara = commentRun?.Parent as Paragraph ?? parent as Paragraph
                    ?? throw new ArgumentException("Comments must be added to a paragraph or run: /body/p[N] or /body/p[N]/r[M]");

                var author = properties.GetValueOrDefault("author", "officecli");
                var initials = properties.GetValueOrDefault("initials", author[..1]);
                var commentsPart = _doc.MainDocumentPart!.WordprocessingCommentsPart
                    ?? _doc.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
                commentsPart.Comments ??= new Comments();

                var commentId = (commentsPart.Comments.Elements<Comment>()
                    .Select(c => int.TryParse(c.Id?.Value, out var id) ? id : 0)
                    .DefaultIfEmpty(0).Max() + 1).ToString();

                commentsPart.Comments.AppendChild(new Comment(
                    new Paragraph(new Run(new Text(commentText) { Space = SpaceProcessingModeValues.Preserve })))
                {
                    Id = commentId, Author = author, Initials = initials,
                    Date = properties.TryGetValue("date", out var ds) ? DateTime.Parse(ds) : DateTime.UtcNow
                });
                commentsPart.Comments.Save();

                var rangeStart = new CommentRangeStart { Id = commentId };
                var rangeEnd = new CommentRangeEnd { Id = commentId };
                var refRun = new Run(new CommentReference { Id = commentId });

                if (commentRun != null)
                {
                    commentRun.InsertBeforeSelf(rangeStart);
                    commentRun.InsertAfterSelf(rangeEnd);
                    rangeEnd.InsertAfterSelf(refRun);
                }
                else
                {
                    var after = commentPara.ParagraphProperties as OpenXmlElement;
                    if (after != null) after.InsertAfterSelf(rangeStart);
                    else commentPara.InsertAt(rangeStart, 0);
                    commentPara.AppendChild(rangeEnd);
                    commentPara.AppendChild(refRun);
                }

                newElement = rangeStart;
                resultPath = $"{parentPath}/comment[{commentId}]";
                break;
            }

            default:
            {
                // Generic fallback: create typed element via SDK schema validation
                var created = GenericXmlQuery.TryCreateTypedElement(parent, type, properties, index);
                if (created == null)
                    throw new ArgumentException($"Schema-invalid element type '{type}' for parent '{parentPath}'. " +
                        "Use raw-set --action append with explicit XML instead.");

                newElement = created;
                var siblings = parent.ChildElements.Where(e => e.LocalName == created.LocalName).ToList();
                var createdIdx = siblings.IndexOf(created) + 1;
                resultPath = $"{parentPath}/{created.LocalName}[{createdIdx}]";
                break;
            }
        }

        _doc.MainDocumentPart?.Document?.Save();
        return resultPath;
    }

    public (string RelId, string PartPath) AddPart(string parentPartPath, string partType, Dictionary<string, string>? properties = null)
    {
        var mainPart = _doc.MainDocumentPart!;

        switch (partType.ToLowerInvariant())
        {
            case "chart":
                var chartPart = mainPart.AddNewPart<ChartPart>();
                var relId = mainPart.GetIdOfPart(chartPart);
                // Initialize with minimal valid ChartSpace
                chartPart.ChartSpace = new C.ChartSpace(
                    new C.Chart(new C.PlotArea(new C.Layout()))
                );
                chartPart.ChartSpace.Save();
                var chartIdx = mainPart.ChartParts.ToList().IndexOf(chartPart);
                return (relId, $"/chart[{chartIdx + 1}]");

            case "header":
                var headerPart = mainPart.AddNewPart<HeaderPart>();
                var hRelId = mainPart.GetIdOfPart(headerPart);
                headerPart.Header = new Header(new Paragraph());
                headerPart.Header.Save();
                var hIdx = mainPart.HeaderParts.ToList().IndexOf(headerPart);
                return (hRelId, $"/header[{hIdx + 1}]");

            case "footer":
                var footerPart = mainPart.AddNewPart<FooterPart>();
                var fRelId = mainPart.GetIdOfPart(footerPart);
                footerPart.Footer = new Footer(new Paragraph());
                footerPart.Footer.Save();
                var fIdx = mainPart.FooterParts.ToList().IndexOf(footerPart);
                return (fRelId, $"/footer[{fIdx + 1}]");

            default:
                throw new ArgumentException(
                    $"Unknown part type: {partType}. Supported: chart, header, footer");
        }
    }

    public void Remove(string path)
    {
        var parts = ParsePath(path);
        var element = NavigateToElement(parts)
            ?? throw new ArgumentException($"Path not found: {path}");

        element.Remove();
        _doc.MainDocumentPart?.Document?.Save();
    }

    public string Move(string sourcePath, string? targetParentPath, int? index)
    {
        var srcParts = ParsePath(sourcePath);
        var element = NavigateToElement(srcParts)
            ?? throw new ArgumentException($"Source not found: {sourcePath}");

        // Determine target parent
        string effectiveParentPath;
        OpenXmlElement targetParent;
        if (string.IsNullOrEmpty(targetParentPath))
        {
            // Reorder within current parent
            targetParent = element.Parent
                ?? throw new InvalidOperationException("Element has no parent");
            // Compute parent path by removing last segment
            var lastSlash = sourcePath.LastIndexOf('/');
            effectiveParentPath = lastSlash > 0 ? sourcePath[..lastSlash] : "/body";
        }
        else
        {
            effectiveParentPath = targetParentPath;
            if (targetParentPath is "/" or "" or "/body")
                targetParent = _doc.MainDocumentPart!.Document!.Body!;
            else
            {
                var tgtParts = ParsePath(targetParentPath);
                targetParent = NavigateToElement(tgtParts)
                    ?? throw new ArgumentException($"Target parent not found: {targetParentPath}");
            }
        }

        element.Remove();
        InsertAtPosition(targetParent, element, index);

        _doc.MainDocumentPart?.Document?.Save();

        var siblings = targetParent.ChildElements.Where(e => e.LocalName == element.LocalName).ToList();
        var newIdx = siblings.IndexOf(element) + 1;
        return $"{effectiveParentPath}/{element.LocalName}[{newIdx}]";
    }

    public string CopyFrom(string sourcePath, string targetParentPath, int? index)
    {
        var srcParts = ParsePath(sourcePath);
        var element = NavigateToElement(srcParts)
            ?? throw new ArgumentException($"Source not found: {sourcePath}");

        var clone = element.CloneNode(true);

        OpenXmlElement targetParent;
        if (targetParentPath is "/" or "" or "/body")
            targetParent = _doc.MainDocumentPart!.Document!.Body!;
        else
        {
            var tgtParts = ParsePath(targetParentPath);
            targetParent = NavigateToElement(tgtParts)
                ?? throw new ArgumentException($"Target parent not found: {targetParentPath}");
        }

        InsertAtPosition(targetParent, clone, index);

        _doc.MainDocumentPart?.Document?.Save();

        var siblings = targetParent.ChildElements.Where(e => e.LocalName == clone.LocalName).ToList();
        var newIdx = siblings.IndexOf(clone) + 1;
        return $"{targetParentPath}/{clone.LocalName}[{newIdx}]";
    }

    private static void InsertAtPosition(OpenXmlElement parent, OpenXmlElement element, int? index)
    {
        if (index.HasValue)
        {
            var children = parent.ChildElements.ToList();
            if (index.Value >= 0 && index.Value < children.Count)
                children[index.Value].InsertBeforeSelf(element);
            else
                parent.AppendChild(element);
        }
        else
        {
            parent.AppendChild(element);
        }
    }

    private void SetDocumentProperties(Dictionary<string, string> properties)
    {
        var doc = _doc.MainDocumentPart?.Document
            ?? throw new InvalidOperationException("Document not found");

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "pagebackground" or "background":
                    doc.DocumentBackground = new DocumentBackground { Color = value };
                    // Enable background display in settings
                    var settingsPart = _doc.MainDocumentPart!.DocumentSettingsPart
                        ?? _doc.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings ??= new Settings();
                    if (settingsPart.Settings.GetFirstChild<DisplayBackgroundShape>() == null)
                        settingsPart.Settings.AppendChild(new DisplayBackgroundShape());
                    settingsPart.Settings.Save();
                    break;

                case "defaultfont":
                    var stylesPart = _doc.MainDocumentPart!.StyleDefinitionsPart;
                    if (stylesPart?.Styles != null)
                    {
                        var defaultRunProps = stylesPart.Styles.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle;
                        if (defaultRunProps != null)
                        {
                            var fonts = defaultRunProps.GetFirstChild<RunFonts>()
                                ?? defaultRunProps.AppendChild(new RunFonts());
                            fonts.Ascii = value;
                            fonts.HighAnsi = value;
                            fonts.EastAsia = value;
                            stylesPart.Styles.Save();
                        }
                    }
                    break;

                case "pagewidth":
                    EnsureSectionProperties().GetFirstChild<PageSize>()!.Width = uint.Parse(value);
                    break;
                case "pageheight":
                    EnsureSectionProperties().GetFirstChild<PageSize>()!.Height = uint.Parse(value);
                    break;
                case "margintop":
                    EnsurePageMargin().Top = int.Parse(value);
                    break;
                case "marginbottom":
                    EnsurePageMargin().Bottom = int.Parse(value);
                    break;
                case "marginleft":
                    EnsurePageMargin().Left = uint.Parse(value);
                    break;
                case "marginright":
                    EnsurePageMargin().Right = uint.Parse(value);
                    break;
            }
        }
    }

    private SectionProperties EnsureSectionProperties()
    {
        var body = _doc.MainDocumentPart!.Document!.Body!;
        var sectPr = body.GetFirstChild<SectionProperties>();
        if (sectPr == null)
        {
            sectPr = new SectionProperties();
            body.AppendChild(sectPr);
        }
        if (sectPr.GetFirstChild<PageSize>() == null)
            sectPr.AppendChild(new PageSize { Width = 11906, Height = 16838 }); // A4 default
        return sectPr;
    }

    private PageMargin EnsurePageMargin()
    {
        var sectPr = EnsureSectionProperties();
        var margin = sectPr.GetFirstChild<PageMargin>();
        if (margin == null)
        {
            margin = new PageMargin { Top = 1440, Bottom = 1440, Left = 1800, Right = 1800 };
            sectPr.AppendChild(margin);
        }
        return margin;
    }

    // ==================== Raw Layer ====================

    public string Raw(string partPath, int? startRow = null, int? endRow = null, HashSet<string>? cols = null)
    {
        var mainPart = _doc.MainDocumentPart;
        if (mainPart == null) return "(no main part)";

        return partPath.ToLowerInvariant() switch
        {
            "/document" or "/word/document.xml" => mainPart.Document?.OuterXml ?? "",
            "/styles" or "/word/styles.xml" => mainPart.StyleDefinitionsPart?.Styles?.OuterXml ?? "(no styles)",
            "/settings" or "/word/settings.xml" => mainPart.DocumentSettingsPart?.Settings?.OuterXml ?? "(no settings)",
            "/numbering" or "/word/numbering.xml" => mainPart.NumberingDefinitionsPart?.Numbering?.OuterXml ?? "(no numbering)",
            "/comments" => mainPart.WordprocessingCommentsPart?.Comments?.OuterXml ?? "(no comments)",
            _ when partPath.StartsWith("/header") => GetHeaderRawXml(partPath),
            _ when partPath.StartsWith("/footer") => GetFooterRawXml(partPath),
            _ when partPath.StartsWith("/chart") => GetChartRawXml(partPath),
            _ => $"Unknown part: {partPath}. Available: /document, /styles, /settings, /numbering, /header[n], /footer[n], /chart[n]"
        };
    }

    public void RawSet(string partPath, string xpath, string action, string? xml)
    {
        var mainPart = _doc.MainDocumentPart
            ?? throw new InvalidOperationException("No main document part");

        OpenXmlPartRootElement rootElement;
        var lowerPath = partPath.ToLowerInvariant();

        if (lowerPath is "/document" or "/")
            rootElement = mainPart.Document ?? throw new InvalidOperationException("No document");
        else if (lowerPath is "/styles")
            rootElement = mainPart.StyleDefinitionsPart?.Styles ?? throw new InvalidOperationException("No styles part");
        else if (lowerPath is "/settings")
            rootElement = mainPart.DocumentSettingsPart?.Settings ?? throw new InvalidOperationException("No settings part");
        else if (lowerPath is "/numbering")
            rootElement = mainPart.NumberingDefinitionsPart?.Numbering ?? throw new InvalidOperationException("No numbering part");
        else if (lowerPath is "/comments")
            rootElement = mainPart.WordprocessingCommentsPart?.Comments ?? throw new InvalidOperationException("No comments part");
        else if (lowerPath.StartsWith("/header"))
        {
            var idx = 0;
            var bracketIdx = partPath.IndexOf('[');
            if (bracketIdx >= 0)
                int.TryParse(partPath[(bracketIdx + 1)..].TrimEnd(']'), out idx);
            var headerPart = mainPart.HeaderParts.ElementAtOrDefault(idx - 1)
                ?? throw new ArgumentException($"header[{idx}] not found");
            rootElement = headerPart.Header ?? throw new InvalidOperationException($"Corrupt file: header[{idx}] data missing");
        }
        else if (lowerPath.StartsWith("/footer"))
        {
            var idx = 0;
            var bracketIdx = partPath.IndexOf('[');
            if (bracketIdx >= 0)
                int.TryParse(partPath[(bracketIdx + 1)..].TrimEnd(']'), out idx);
            var footerPart = mainPart.FooterParts.ElementAtOrDefault(idx - 1)
                ?? throw new ArgumentException($"footer[{idx}] not found");
            rootElement = footerPart.Footer ?? throw new InvalidOperationException($"Corrupt file: footer[{idx}] data missing");
        }
        else if (lowerPath.StartsWith("/chart"))
        {
            var idx = 0;
            var bracketIdx = partPath.IndexOf('[');
            if (bracketIdx >= 0)
                int.TryParse(partPath[(bracketIdx + 1)..].TrimEnd(']'), out idx);
            var chartPart = mainPart.ChartParts.ElementAtOrDefault(idx - 1)
                ?? throw new ArgumentException($"chart[{idx}] not found");
            rootElement = chartPart.ChartSpace ?? throw new InvalidOperationException($"Corrupt file: chart[{idx}] data missing");
        }
        else
            throw new ArgumentException($"Unknown part: {partPath}. Available: /document, /styles, /settings, /numbering, /header[n], /footer[n], /chart[n]");

        var affected = RawXmlHelper.Execute(rootElement, xpath, action, xml);
        rootElement.Save();
        Console.WriteLine($"raw-set: {affected} element(s) affected");
    }

    public List<ValidationError> Validate() => RawXmlHelper.ValidateDocument(_doc);

    public void Dispose()
    {
        _doc.Dispose();
    }

    // ==================== Private Helpers ====================

    private static string GetParagraphText(Paragraph para)
    {
        return string.Concat(para.Descendants<Text>().Select(t => t.Text));
    }

    /// <summary>
    /// Get paragraph text including inline math rendered as readable Unicode.
    /// </summary>
    private static string GetParagraphTextWithMath(Paragraph para)
    {
        var sb = new StringBuilder();
        foreach (var child in para.ChildElements)
        {
            if (child is Run run)
                sb.Append(GetRunText(run));
            else if (child.LocalName == "oMath" || child is M.OfficeMath)
                sb.Append(FormulaParser.ToReadableText(child));
            else if (child is Hyperlink hyperlink)
                sb.Append(string.Concat(hyperlink.Descendants<Text>().Select(t => t.Text)));
        }
        return sb.ToString();
    }

    /// <summary>
    /// Find math elements in a paragraph using both type and localName matching.
    /// </summary>
    private static List<OpenXmlElement> FindMathElements(Paragraph para)
    {
        return para.ChildElements
            .Where(e => e.LocalName == "oMath" || e is M.OfficeMath)
            .ToList();
    }

    /// <summary>
    /// Get all body-level elements, flattening SdtContent containers.
    /// This ensures paragraphs and tables inside w:sdt are not missed.
    /// </summary>
    private static IEnumerable<OpenXmlElement> GetBodyElements(Body body)
    {
        foreach (var element in body.ChildElements)
        {
            if (element is SdtBlock sdt)
            {
                var content = sdt.SdtContentBlock;
                if (content != null)
                {
                    foreach (var child in content.ChildElements)
                        yield return child;
                }
            }
            else
            {
                yield return element;
            }
        }
    }

    /// <summary>
    /// Checks if an element is a structural document element worth displaying
    /// (not inline markers like bookmarkStart, bookmarkEnd, proofErr, etc.)
    /// </summary>
    private static bool IsStructuralElement(OpenXmlElement element)
    {
        var name = element.LocalName;
        return name == "sectPr" || name == "altChunk" || name == "customXml";
    }

    /// <summary>
    /// Get all Run elements in a paragraph, including those nested inside
    /// Hyperlink and SdtContent containers.
    /// </summary>
    private static List<Run> GetAllRuns(Paragraph para)
    {
        return para.Descendants<Run>().ToList();
    }

    private static string GetRunText(Run run)
    {
        return string.Concat(run.Elements<Text>().Select(t => t.Text));
    }

    private string GetStyleName(Paragraph para)
    {
        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (styleId == null) return "Normal";

        // Try to resolve display name from styles part
        var stylesPart = _doc.MainDocumentPart?.StyleDefinitionsPart;
        if (stylesPart?.Styles != null)
        {
            var style = stylesPart.Styles.Elements<Style>()
                .FirstOrDefault(s => s.StyleId?.Value == styleId);
            if (style?.StyleName?.Val?.Value != null)
                return style.StyleName.Val.Value;
        }

        return styleId;
    }

    private static string? GetRunFont(Run run)
    {
        var fonts = run.RunProperties?.RunFonts;
        return fonts?.EastAsia?.Value ?? fonts?.Ascii?.Value ?? fonts?.HighAnsi?.Value;
    }

    private static string? GetRunFontSize(Run run)
    {
        var size = run.RunProperties?.FontSize?.Val?.Value;
        if (size == null) return null;
        return $"{int.Parse(size) / 2}pt"; // stored as half-points
    }

    private string GetRunFormatDescription(Run run, Paragraph? para = null)
    {
        var parts = new List<string>();

        RunProperties? rProps;
        if (para != null)
        {
            rProps = ResolveEffectiveRunProperties(run, para);
        }
        else
        {
            rProps = run.RunProperties;
        }
        if (rProps == null) return "(default)";

        var font = GetFontFromProperties(rProps);
        if (font != null) parts.Add(font);

        var size = GetSizeFromProperties(rProps);
        if (size != null) parts.Add(size);

        if (rProps.Bold != null) parts.Add("bold");
        if (rProps.Italic != null) parts.Add("italic");
        if (rProps.Underline != null) parts.Add("underline");
        if (rProps.Strike != null) parts.Add("strikethrough");

        return parts.Count > 0 ? string.Join(" ", parts) : "(default)";
    }

    private static int GetHeadingLevel(string styleName)
    {
        // Heading 1, Heading 2, heading1, 标题 1, etc.
        foreach (var ch in styleName)
        {
            if (char.IsDigit(ch))
                return ch - '0';
        }
        if (styleName == "Title") return 0;
        if (styleName == "Subtitle") return 1;
        return 1;
    }

    private static bool IsNormalStyle(string styleName)
    {
        return styleName is "Normal" or "正文" or "Body Text" or "Body" or "a"
            || styleName.StartsWith("Normal");
    }

    private string? FindWatermark()
    {
        var headerParts = _doc.MainDocumentPart?.HeaderParts;
        if (headerParts == null) return null;

        foreach (var headerPart in headerParts)
        {
            var header = headerPart.Header;
            if (header == null) continue;

            // Search for VML shapes with watermark
            foreach (var pict in header.Descendants<DocumentFormat.OpenXml.Wordprocessing.Picture>())
            {
                foreach (var shape in pict.Descendants<Vml.Shape>())
                {
                    var id = shape.GetAttribute("id", "");
                    if (id.Value?.Contains("WaterMark", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        var textPath = shape.Descendants<Vml.TextPath>().FirstOrDefault();
                        return textPath?.String?.Value ?? "(image watermark)";
                    }
                }
            }

            // Also check for DrawingML watermarks
            foreach (var drawing in header.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>())
            {
                // Simple detection: check if it looks like a watermark by inline/anchor properties
                var docProps = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>().FirstOrDefault();
                if (docProps?.Name?.Value?.Contains("WaterMark", StringComparison.OrdinalIgnoreCase) == true)
                {
                    return "(image watermark)";
                }
            }
        }

        return null;
    }

    private List<string> GetHeaderTexts()
    {
        var results = new List<string>();
        var headerParts = _doc.MainDocumentPart?.HeaderParts;
        if (headerParts == null) return results;

        foreach (var headerPart in headerParts)
        {
            var header = headerPart.Header;
            if (header == null) continue;
            var text = string.Concat(header.Descendants<Text>().Select(t => t.Text)).Trim();
            if (!string.IsNullOrEmpty(text))
                results.Add(text);
        }

        return results;
    }

    private List<string> GetFooterTexts()
    {
        var results = new List<string>();
        var footerParts = _doc.MainDocumentPart?.FooterParts;
        if (footerParts == null) return results;

        foreach (var footerPart in footerParts)
        {
            var footer = footerPart.Footer;
            if (footer == null) continue;
            var text = string.Concat(footer.Descendants<Text>().Select(t => t.Text)).Trim();
            if (!string.IsNullOrEmpty(text))
                results.Add(text);
            else
            {
                // Check for page numbers
                var fldChars = footer.Descendants<FieldCode>().Any();
                if (fldChars)
                    results.Add("(page number)");
            }
        }

        return results;
    }

    private static bool HasMixedPunctuation(string text)
    {
        var chinesePunct = "\uff0c\u3002\uff01\uff1f\u3001\uff1b\uff1a\u201c\u201d\u2018\u2019\uff08\uff09\u3010\u3011";
        bool hasChinese = text.Any(c => chinesePunct.Contains(c));
        bool hasEnglish = text.Any(c => ",.!?;:\"'()[]".Contains(c));
        bool hasChineseChars = text.Any(c => c >= 0x4E00 && c <= 0x9FFF);
        return hasChinese && hasEnglish && hasChineseChars;
    }

    private static RunProperties EnsureRunProperties(Run run)
    {
        return run.RunProperties ?? run.PrependChild(new RunProperties());
    }

    // ==================== Navigation ====================

    private DocumentNode GetRootNode(int depth)
    {
        var node = new DocumentNode { Path = "/", Type = "document" };
        var children = new List<DocumentNode>();

        var mainPart = _doc.MainDocumentPart;
        if (mainPart?.Document?.Body != null)
        {
            children.Add(new DocumentNode
            {
                Path = "/body",
                Type = "body",
                ChildCount = mainPart.Document.Body.ChildElements.Count
            });
        }

        if (mainPart?.StyleDefinitionsPart != null)
        {
            children.Add(new DocumentNode
            {
                Path = "/styles",
                Type = "styles",
                ChildCount = mainPart.StyleDefinitionsPart.Styles?.ChildElements.Count ?? 0
            });
        }

        int headerIdx = 0;
        if (mainPart?.HeaderParts != null)
        {
            foreach (var _ in mainPart.HeaderParts)
            {
                children.Add(new DocumentNode
                {
                    Path = $"/header[{headerIdx + 1}]",
                    Type = "header"
                });
                headerIdx++;
            }
        }

        int footerIdx = 0;
        if (mainPart?.FooterParts != null)
        {
            foreach (var _ in mainPart.FooterParts)
            {
                children.Add(new DocumentNode
                {
                    Path = $"/footer[{footerIdx + 1}]",
                    Type = "footer"
                });
                footerIdx++;
            }
        }

        if (mainPart?.NumberingDefinitionsPart != null)
        {
            children.Add(new DocumentNode { Path = "/numbering", Type = "numbering" });
        }

        node.Children = children;
        node.ChildCount = children.Count;
        return node;
    }

    private record PathSegment(string Name, int? Index);

    private static List<PathSegment> ParsePath(string path)
    {
        var segments = new List<PathSegment>();
        var parts = path.Trim('/').Split('/');

        foreach (var part in parts)
        {
            var bracketIdx = part.IndexOf('[');
            if (bracketIdx >= 0)
            {
                var name = part[..bracketIdx];
                var indexStr = part[(bracketIdx + 1)..^1];
                segments.Add(new PathSegment(name, int.Parse(indexStr)));
            }
            else
            {
                segments.Add(new PathSegment(part, null));
            }
        }

        return segments;
    }

    private OpenXmlElement? NavigateToElement(List<PathSegment> segments)
    {
        if (segments.Count == 0) return null;

        var first = segments[0];
        OpenXmlElement? current = first.Name.ToLowerInvariant() switch
        {
            "body" => _doc.MainDocumentPart?.Document?.Body,
            "styles" => _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles,
            "header" => _doc.MainDocumentPart?.HeaderParts.ElementAtOrDefault((first.Index ?? 1) - 1)?.Header,
            "footer" => _doc.MainDocumentPart?.FooterParts.ElementAtOrDefault((first.Index ?? 1) - 1)?.Footer,
            "numbering" => _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering,
            "settings" => _doc.MainDocumentPart?.DocumentSettingsPart?.Settings,
            "comments" => _doc.MainDocumentPart?.WordprocessingCommentsPart?.Comments,
            _ => null
        };

        for (int i = 1; i < segments.Count && current != null; i++)
        {
            var seg = segments[i];
            IEnumerable<OpenXmlElement> children;
            if (current is Body body2 && (seg.Name.ToLowerInvariant() == "p" || seg.Name.ToLowerInvariant() == "tbl"))
            {
                // Flatten sdt containers when navigating body-level paragraphs/tables
                children = seg.Name.ToLowerInvariant() == "p"
                    ? GetBodyElements(body2).OfType<Paragraph>().Cast<OpenXmlElement>()
                    : GetBodyElements(body2).OfType<Table>().Cast<OpenXmlElement>();
            }
            else
            {
                children = seg.Name.ToLowerInvariant() switch
                {
                    "p" => current.Elements<Paragraph>().Cast<OpenXmlElement>(),
                    "r" => current.Descendants<Run>()
                        .Where(r => r.GetFirstChild<CommentReference>() == null)
                        .Cast<OpenXmlElement>(),
                    "tbl" => current.Elements<Table>().Cast<OpenXmlElement>(),
                    "tr" => current.Elements<TableRow>().Cast<OpenXmlElement>(),
                    "tc" => current.Elements<TableCell>().Cast<OpenXmlElement>(),
                    _ => current.ChildElements.Where(e => e.LocalName == seg.Name).Cast<OpenXmlElement>()
                };
            }

            current = seg.Index.HasValue
                ? children.ElementAtOrDefault(seg.Index.Value - 1)
                : children.FirstOrDefault();
        }

        return current;
    }

    private DocumentNode ElementToNode(OpenXmlElement element, string path, int depth)
    {
        var node = new DocumentNode { Path = path, Type = element.LocalName };

        if (element is Paragraph para)
        {
            node.Type = "paragraph";
            node.Text = GetParagraphText(para);
            node.Style = GetStyleName(para);
            node.Preview = node.Text?.Length > 50 ? node.Text[..50] + "..." : node.Text;
            node.ChildCount = GetAllRuns(para).Count();

            var pProps = para.ParagraphProperties;
            if (pProps != null)
            {
                if (pProps.Justification?.Val?.Value != null)
                    node.Format["alignment"] = pProps.Justification.Val.Value.ToString();
                if (pProps.Indentation?.FirstLine?.Value != null)
                    node.Format["firstLineIndent"] = pProps.Indentation.FirstLine.Value;
            }

            if (depth > 0)
            {
                int runIdx = 0;
                foreach (var run in GetAllRuns(para))
                {
                    node.Children.Add(ElementToNode(run, $"{path}/r[{runIdx + 1}]", depth - 1));
                    runIdx++;
                }
            }
        }
        else if (element is Run run)
        {
            node.Type = "run";
            node.Text = GetRunText(run);
            var font = GetRunFont(run);
            if (font != null) node.Format["font"] = font;
            var size = GetRunFontSize(run);
            if (size != null) node.Format["size"] = size;
            if (run.RunProperties?.Bold != null) node.Format["bold"] = true;
            if (run.RunProperties?.Italic != null) node.Format["italic"] = true;
        }
        else if (element is Table table)
        {
            node.Type = "table";
            node.ChildCount = table.Elements<TableRow>().Count();
            var firstRow = table.Elements<TableRow>().FirstOrDefault();
            node.Format["cols"] = firstRow?.Elements<TableCell>().Count() ?? 0;

            if (depth > 0)
            {
                int rowIdx = 0;
                foreach (var row in table.Elements<TableRow>())
                {
                    var rowNode = new DocumentNode
                    {
                        Path = $"{path}/tr[{rowIdx + 1}]",
                        Type = "row",
                        ChildCount = row.Elements<TableCell>().Count()
                    };
                    if (depth > 1)
                    {
                        int cellIdx = 0;
                        foreach (var cell in row.Elements<TableCell>())
                        {
                            var cellNode = new DocumentNode
                            {
                                Path = $"{path}/tr[{rowIdx + 1}]/tc[{cellIdx + 1}]",
                                Type = "cell",
                                Text = string.Join("", cell.Descendants<Text>().Select(t => t.Text)),
                                ChildCount = cell.Elements<Paragraph>().Count()
                            };
                            if (depth > 2)
                            {
                                int pIdx = 0;
                                foreach (var cellPara in cell.Elements<Paragraph>())
                                {
                                    cellNode.Children.Add(ElementToNode(cellPara, $"{path}/tr[{rowIdx + 1}]/tc[{cellIdx + 1}]/p[{pIdx + 1}]", depth - 3));
                                    pIdx++;
                                }
                            }
                            rowNode.Children.Add(cellNode);
                            cellIdx++;
                        }
                    }
                    node.Children.Add(rowNode);
                    rowIdx++;
                }
            }
        }
        else
        {
            // Generic fallback: collect XML attributes and child val patterns
            foreach (var attr in element.GetAttributes())
                node.Format[attr.LocalName] = attr.Value;
            foreach (var child in element.ChildElements)
            {
                if (child.ChildElements.Count == 0)
                {
                    foreach (var attr in child.GetAttributes())
                    {
                        if (attr.LocalName.Equals("val", StringComparison.OrdinalIgnoreCase))
                        {
                            node.Format[child.LocalName] = attr.Value;
                            break;
                        }
                    }
                }
            }

            var innerText = element.InnerText;
            if (!string.IsNullOrEmpty(innerText))
                node.Text = innerText.Length > 200 ? innerText[..200] + "..." : innerText;
            if (string.IsNullOrEmpty(innerText))
            {
                var outerXml = element.OuterXml;
                node.Preview = outerXml.Length > 200 ? outerXml[..200] + "..." : outerXml;
            }

            node.ChildCount = element.ChildElements.Count;
            if (depth > 0)
            {
                var typeCounters = new Dictionary<string, int>();
                foreach (var child in element.ChildElements)
                {
                    var name = child.LocalName;
                    typeCounters.TryGetValue(name, out int idx);
                    node.Children.Add(ElementToNode(child, $"{path}/{name}[{idx + 1}]", depth - 1));
                    typeCounters[name] = idx + 1;
                }
            }
        }

        return node;
    }

    // ==================== Selector ====================

    private record SelectorPart(string? Element, Dictionary<string, string> Attributes, string? ContainsText, SelectorPart? ChildSelector);

    private static SelectorPart ParseSelector(string selector)
    {
        // Support: element[attr=value] > child[attr=value]
        var childParts = selector.Split('>').Select(s => s.Trim()).ToArray();

        SelectorPart? childSelector = null;
        if (childParts.Length > 1)
        {
            childSelector = ParseSingleSelector(childParts[1]);
        }

        var main = ParseSingleSelector(childParts[0]);
        return main with { ChildSelector = childSelector };
    }

    private static SelectorPart ParseSingleSelector(string selector)
    {
        var attrs = new Dictionary<string, string>();
        string? element = null;
        string? containsText = null;

        // Extract element name (before any [ or : modifier)
        var firstMod = selector.Length;
        var bracketIdx = selector.IndexOf('[');
        if (bracketIdx >= 0 && bracketIdx < firstMod) firstMod = bracketIdx;
        var colonIdx = selector.IndexOf(':');
        if (colonIdx >= 0 && colonIdx < firstMod) firstMod = colonIdx;

        element = selector[..firstMod].Trim();
        if (string.IsNullOrEmpty(element)) element = null;

        // Parse [attr=value] attributes
        if (bracketIdx >= 0)
        {
            var attrPart = selector[bracketIdx..];
            foreach (System.Text.RegularExpressions.Match m in
                System.Text.RegularExpressions.Regex.Matches(attrPart, @"\[(\w+)(!?=)([^\]]+)\]"))
            {
                var key = m.Groups[1].Value;
                var op = m.Groups[2].Value;
                var val = m.Groups[3].Value.Trim('\'', '"');
                attrs[key] = (op == "!=" ? "!" : "") + val;
            }
        }

        // Parse :contains("text") pseudo-selector
        if (selector.Contains(":contains("))
        {
            var idx = selector.IndexOf(":contains(");
            var endIdx = selector.IndexOf(')', idx + 10);
            if (endIdx >= 0)
                containsText = selector[(idx + 10)..endIdx].Trim('\'', '"');
        }

        // Parse :empty pseudo-selector
        if (selector.Contains(":empty"))
        {
            attrs["__empty"] = "true";
        }

        // Parse :no-alt pseudo-selector
        if (selector.Contains(":no-alt"))
        {
            attrs["__no-alt"] = "true";
        }

        return new SelectorPart(element, attrs, containsText, null);
    }

    private bool MatchesSelector(Paragraph para, SelectorPart selector, int lineNum)
    {
        // If selector targets runs (has child selector), only match parent paragraph
        if (selector.ChildSelector != null)
        {
            // Check paragraph-level attributes
            if (selector.Element != null && selector.Element != "p" && selector.Element != "paragraph")
                return false;
            return MatchesParagraphAttrs(para, selector.Attributes);
        }

        if (selector.Element != null && selector.Element != "p" && selector.Element != "paragraph")
            return false;

        if (!MatchesParagraphAttrs(para, selector.Attributes))
            return false;

        if (selector.Attributes.ContainsKey("__empty"))
        {
            return string.IsNullOrWhiteSpace(GetParagraphText(para));
        }

        if (selector.ContainsText != null)
        {
            return GetParagraphText(para).Contains(selector.ContainsText);
        }

        return true;
    }

    private bool MatchesParagraphAttrs(Paragraph para, Dictionary<string, string> attrs)
    {
        foreach (var (key, rawVal) in attrs)
        {
            if (key == "__empty") continue;
            bool negate = rawVal.StartsWith("!");
            var val = negate ? rawVal[1..] : rawVal;

            string? actual = key.ToLowerInvariant() switch
            {
                "style" => GetStyleName(para),
                "alignment" => para.ParagraphProperties?.Justification?.Val?.HasValue == true
                    ? para.ParagraphProperties.Justification.Val.Value.ToString() : null,
                "firstlineindent" => para.ParagraphProperties?.Indentation?.FirstLine?.Value,
                _ => GenericXmlQuery.GetAttributeValue(para, key)
                     ?? (para.ParagraphProperties != null ? GenericXmlQuery.GetAttributeValue(para.ParagraphProperties, key) : null)
            };

            // For style, also match against styleId (e.g., "Heading1" vs display name "heading 1")
            bool matches;
            if (key.Equals("style", StringComparison.OrdinalIgnoreCase))
            {
                var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                matches = string.Equals(actual, val, StringComparison.OrdinalIgnoreCase)
                       || string.Equals(styleId, val, StringComparison.OrdinalIgnoreCase);
            }
            else
            {
                matches = string.Equals(actual, val, StringComparison.OrdinalIgnoreCase);
            }
            if (negate ? matches : !matches) return false;
        }
        return true;
    }

    private static bool MatchesRunSelector(Run run, Paragraph parent, SelectorPart selector)
    {
        if (selector.Element != null && selector.Element != "r" && selector.Element != "run")
            return false;

        foreach (var (key, rawVal) in selector.Attributes)
        {
            bool negate = rawVal.StartsWith("!");
            var val = negate ? rawVal[1..] : rawVal;

            string? actual = key.ToLowerInvariant() switch
            {
                "font" => GetRunFont(run),
                "size" => GetRunFontSize(run),
                "bold" => run.RunProperties?.Bold != null ? "true" : "false",
                "italic" => run.RunProperties?.Italic != null ? "true" : "false",
                _ => GenericXmlQuery.GetAttributeValue(run, key)
                     ?? (run.RunProperties != null ? GenericXmlQuery.GetAttributeValue(run.RunProperties, key) : null)
            };

            bool matches = string.Equals(actual, val, StringComparison.OrdinalIgnoreCase);
            if (negate ? matches : !matches) return false;
        }

        if (selector.ContainsText != null)
        {
            return GetRunText(run).Contains(selector.ContainsText);
        }

        return true;
    }

    private string GetHeaderRawXml(string partPath)
    {
        var idx = 0;
        var bracketIdx = partPath.IndexOf('[');
        if (bracketIdx >= 0)
            int.TryParse(partPath[(bracketIdx + 1)..^0].TrimEnd(']'), out idx);

        var headerPart = _doc.MainDocumentPart?.HeaderParts.ElementAtOrDefault(idx);
        return headerPart?.Header?.OuterXml ?? $"(header[{idx}] not found)";
    }

    private string GetFooterRawXml(string partPath)
    {
        var idx = 0;
        var bracketIdx = partPath.IndexOf('[');
        if (bracketIdx >= 0)
            int.TryParse(partPath[(bracketIdx + 1)..^0].TrimEnd(']'), out idx);

        var footerPart = _doc.MainDocumentPart?.FooterParts.ElementAtOrDefault(idx);
        return footerPart?.Footer?.OuterXml ?? $"(footer[{idx}] not found)";
    }

    private string GetChartRawXml(string partPath)
    {
        var idx = 0;
        var bracketIdx = partPath.IndexOf('[');
        if (bracketIdx >= 0)
            int.TryParse(partPath[(bracketIdx + 1)..^0].TrimEnd(']'), out idx);

        var chartPart = _doc.MainDocumentPart?.ChartParts.ElementAtOrDefault(idx);
        return chartPart?.ChartSpace?.OuterXml ?? $"(chart[{idx}] not found)";
    }

    // ==================== Style Inheritance ====================

    private RunProperties ResolveEffectiveRunProperties(Run run, Paragraph para)
    {
        var effective = new RunProperties();

        // 1. Start with docDefaults rPr
        var docDefaults = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults;
        var defaultRPr = docDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle;
        if (defaultRPr != null)
            MergeRunProperties(effective, defaultRPr);

        // 2. Walk paragraph style basedOn chain (collect in order, apply from base to derived)
        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (styleId != null)
        {
            var chain = new List<Style>();
            var visited = new HashSet<string>();
            var currentStyleId = styleId;
            while (currentStyleId != null && visited.Add(currentStyleId))
            {
                var style = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                    ?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == currentStyleId);
                if (style == null) break;
                chain.Add(style);
                currentStyleId = style.BasedOn?.Val?.Value;
            }
            // Apply from base to derived (reverse order)
            for (int i = chain.Count - 1; i >= 0; i--)
            {
                var styleRPr = chain[i].StyleRunProperties;
                if (styleRPr != null)
                    MergeRunProperties(effective, styleRPr);
            }
        }

        // 3. Apply run's own rPr (highest priority)
        if (run.RunProperties != null)
            MergeRunProperties(effective, run.RunProperties);

        return effective;
    }

    private static void MergeRunProperties(RunProperties target, OpenXmlElement source)
    {
        // Merge each known property: source overwrites target
        var srcFonts = source.GetFirstChild<RunFonts>();
        if (srcFonts != null)
            target.RunFonts = srcFonts.CloneNode(true) as RunFonts;

        var srcSize = source.GetFirstChild<FontSize>();
        if (srcSize != null)
            target.FontSize = srcSize.CloneNode(true) as FontSize;

        var srcBold = source.GetFirstChild<Bold>();
        if (srcBold != null)
            target.Bold = srcBold.CloneNode(true) as Bold;

        var srcItalic = source.GetFirstChild<Italic>();
        if (srcItalic != null)
            target.Italic = srcItalic.CloneNode(true) as Italic;

        var srcUnderline = source.GetFirstChild<Underline>();
        if (srcUnderline != null)
            target.Underline = srcUnderline.CloneNode(true) as Underline;

        var srcStrike = source.GetFirstChild<Strike>();
        if (srcStrike != null)
            target.Strike = srcStrike.CloneNode(true) as Strike;

        var srcColor = source.GetFirstChild<Color>();
        if (srcColor != null)
            target.Color = srcColor.CloneNode(true) as Color;

        var srcHighlight = source.GetFirstChild<Highlight>();
        if (srcHighlight != null)
            target.Highlight = srcHighlight.CloneNode(true) as Highlight;
    }

    private static string? GetFontFromProperties(RunProperties? rProps)
    {
        if (rProps == null) return null;
        var fonts = rProps.RunFonts;
        return fonts?.EastAsia?.Value ?? fonts?.Ascii?.Value ?? fonts?.HighAnsi?.Value;
    }

    private static string? GetSizeFromProperties(RunProperties? rProps)
    {
        if (rProps == null) return null;
        var size = rProps.FontSize?.Val?.Value;
        if (size == null) return null;
        return $"{int.Parse(size) / 2}pt";
    }

    // ==================== List / Numbering ====================

    private string GetListPrefix(Paragraph para)
    {
        var numProps = para.ParagraphProperties?.NumberingProperties;
        if (numProps == null) return "";

        var numId = numProps.NumberingId?.Val?.Value;
        var ilvl = numProps.NumberingLevelReference?.Val?.Value ?? 0;
        if (numId == null || numId == 0) return "";

        var indent = new string(' ', ilvl * 2);
        var numFmt = GetNumberingFormat(numId.Value, ilvl);

        return numFmt.ToLowerInvariant() switch
        {
            "bullet" => $"{indent}• ",
            "decimal" => $"{indent}1. ",
            "lowerletter" => $"{indent}a. ",
            "upperletter" => $"{indent}A. ",
            "lowerroman" => $"{indent}i. ",
            "upperroman" => $"{indent}I. ",
            _ => $"{indent}• "
        };
    }

    private string GetNumberingFormat(int numId, int ilvl)
    {
        var numbering = _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
        if (numbering == null) return "bullet";

        var numInstance = numbering.Elements<NumberingInstance>()
            .FirstOrDefault(n => n.NumberID?.Value == numId);
        if (numInstance == null) return "bullet";

        var abstractNumId = numInstance.AbstractNumId?.Val?.Value;
        if (abstractNumId == null) return "bullet";

        var abstractNum = numbering.Elements<AbstractNum>()
            .FirstOrDefault(a => a.AbstractNumberId?.Value == abstractNumId);
        if (abstractNum == null) return "bullet";

        var level = abstractNum.Elements<Level>()
            .FirstOrDefault(l => l.LevelIndex?.Value == ilvl);

        var numFmt = level?.NumberingFormat?.Val;
        if (numFmt == null || !numFmt.HasValue) return "bullet";
        return numFmt.InnerText ?? "bullet";
    }

    private void ApplyListStyle(Paragraph para, string listStyleValue)
    {
        var mainPart = _doc.MainDocumentPart!;
        var numberingPart = mainPart.NumberingDefinitionsPart;
        if (numberingPart == null)
        {
            numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering = new Numbering();
        }
        var numbering = numberingPart.Numbering
            ?? throw new InvalidOperationException("Corrupt file: numbering data missing");

        // Determine the next available IDs
        var maxAbstractId = numbering.Elements<AbstractNum>()
            .Select(a => a.AbstractNumberId?.Value ?? 0).DefaultIfEmpty(-1).Max() + 1;
        var maxNumId = numbering.Elements<NumberingInstance>()
            .Select(n => n.NumberID?.Value ?? 0).DefaultIfEmpty(0).Max() + 1;

        var isBullet = listStyleValue.ToLowerInvariant() is "bullet" or "unordered" or "ul";

        // Create abstract numbering definition
        var abstractNum = new AbstractNum { AbstractNumberId = maxAbstractId };
        abstractNum.AppendChild(new MultiLevelType { Val = MultiLevelValues.HybridMultilevel });

        var bulletChars = new[] { "\u2022", "\u25E6", "\u25AA" }; // •, ◦, ▪

        for (int lvl = 0; lvl < 3; lvl++)
        {
            var level = new Level { LevelIndex = lvl };
            level.AppendChild(new StartNumberingValue { Val = 1 });

            if (isBullet)
            {
                level.AppendChild(new NumberingFormat { Val = NumberFormatValues.Bullet });
                level.AppendChild(new LevelText { Val = bulletChars[lvl % bulletChars.Length] });
            }
            else
            {
                var fmt = lvl switch
                {
                    0 => NumberFormatValues.Decimal,
                    1 => NumberFormatValues.LowerLetter,
                    _ => NumberFormatValues.LowerRoman
                };
                level.AppendChild(new NumberingFormat { Val = fmt });
                level.AppendChild(new LevelText { Val = $"%{lvl + 1}." });
            }

            level.AppendChild(new LevelJustification { Val = LevelJustificationValues.Left });
            level.AppendChild(new PreviousParagraphProperties(
                new Indentation { Left = ((lvl + 1) * 720).ToString(), Hanging = "360" }
            ));
            abstractNum.AppendChild(level);
        }

        // Insert AbstractNum before any NumberingInstance elements
        var firstNumInstance = numbering.GetFirstChild<NumberingInstance>();
        if (firstNumInstance != null)
            numbering.InsertBefore(abstractNum, firstNumInstance);
        else
            numbering.AppendChild(abstractNum);

        // Create numbering instance
        var numInstance = new NumberingInstance { NumberID = maxNumId };
        numInstance.AppendChild(new AbstractNumId { Val = maxAbstractId });
        numbering.AppendChild(numInstance);

        numbering.Save();

        // Apply to paragraph
        var pProps = para.ParagraphProperties ?? para.PrependChild(new ParagraphProperties());
        pProps.NumberingProperties = new NumberingProperties
        {
            NumberingId = new NumberingId { Val = maxNumId },
            NumberingLevelReference = new NumberingLevelReference { Val = 0 }
        };
    }

    // ==================== Image Helpers ====================

    private static long ParseEmu(string value)
    {
        // Support: raw EMU number, or suffixed with cm/in/pt/px
        value = value.Trim();
        if (value.EndsWith("cm", StringComparison.OrdinalIgnoreCase))
            return (long)(double.Parse(value[..^2]) * 360000);
        if (value.EndsWith("in", StringComparison.OrdinalIgnoreCase))
            return (long)(double.Parse(value[..^2]) * 914400);
        if (value.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
            return (long)(double.Parse(value[..^2]) * 12700);
        if (value.EndsWith("px", StringComparison.OrdinalIgnoreCase))
            return (long)(double.Parse(value[..^2]) * 9525);
        return long.Parse(value); // raw EMU
    }

    private static Run CreateImageRun(string relationshipId, long cx, long cy, string altText)
    {
        var inline = new DW.Inline(
            new DW.Extent { Cx = cx, Cy = cy },
            new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
            new DW.DocProperties { Id = (uint)Environment.TickCount, Name = altText, Description = altText },
            new DW.NonVisualGraphicFrameDrawingProperties(
                new A.GraphicFrameLocks { NoChangeAspect = true }
            ),
            new A.Graphic(
                new A.GraphicData(
                    new PIC.Picture(
                        new PIC.NonVisualPictureProperties(
                            new PIC.NonVisualDrawingProperties { Id = 0U, Name = altText },
                            new PIC.NonVisualPictureDrawingProperties()
                        ),
                        new PIC.BlipFill(
                            new A.Blip { Embed = relationshipId, CompressionState = A.BlipCompressionValues.Print },
                            new A.Stretch(new A.FillRectangle())
                        ),
                        new PIC.ShapeProperties(
                            new A.Transform2D(
                                new A.Offset { X = 0L, Y = 0L },
                                new A.Extents { Cx = cx, Cy = cy }
                            ),
                            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                        )
                    )
                ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
            )
        )
        {
            DistanceFromTop = 0U,
            DistanceFromBottom = 0U,
            DistanceFromLeft = 0U,
            DistanceFromRight = 0U
        };

        return new Run(new Drawing(inline));
    }

    private static string GetDrawingInfo(Drawing drawing)
    {
        var docProps = drawing.Descendants<DW.DocProperties>().FirstOrDefault();
        var extent = drawing.Descendants<DW.Extent>().FirstOrDefault();

        var parts = new List<string>();
        if (docProps?.Description?.Value is string desc && !string.IsNullOrEmpty(desc))
            parts.Add($"alt=\"{desc}\"");
        else if (docProps?.Name?.Value is string name && !string.IsNullOrEmpty(name))
            parts.Add($"name=\"{name}\"");
        if (extent != null)
        {
            var wCm = extent.Cx != null ? $"{extent.Cx.Value / 360000.0:F1}cm" : "?";
            var hCm = extent.Cy != null ? $"{extent.Cy.Value / 360000.0:F1}cm" : "?";
            parts.Add($"{wCm}×{hCm}");
        }
        return parts.Count > 0 ? string.Join(", ", parts) : "unknown";
    }

    private static DocumentNode CreateImageNode(Drawing drawing, Run run, string path)
    {
        var docProps = drawing.Descendants<DW.DocProperties>().FirstOrDefault();
        var extent = drawing.Descendants<DW.Extent>().FirstOrDefault();

        var node = new DocumentNode
        {
            Path = path,
            Type = "picture",
            Text = docProps?.Description?.Value ?? docProps?.Name?.Value ?? ""
        };
        if (extent?.Cx != null) node.Format["width"] = $"{extent.Cx.Value / 360000.0:F1}cm";
        if (extent?.Cy != null) node.Format["height"] = $"{extent.Cy.Value / 360000.0:F1}cm";
        if (docProps?.Description?.Value != null) node.Format["alt"] = docProps.Description.Value;

        return node;
    }
}
