// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public class PowerPointHandler : IDocumentHandler
{
    private readonly PresentationDocument _doc;
    private readonly string _filePath;

    public PowerPointHandler(string filePath, bool editable)
    {
        _filePath = filePath;
        _doc = PresentationDocument.Open(filePath, editable);
    }

    public string ViewAsText(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
    {
        var sb = new StringBuilder();
        int slideNum = 0;
        int totalSlides = GetSlideParts().Count();

        foreach (var slidePart in GetSlideParts())
        {
            slideNum++;
            if (startLine.HasValue && slideNum < startLine.Value) continue;
            if (endLine.HasValue && slideNum > endLine.Value) break;

            if (maxLines.HasValue && slideNum - (startLine ?? 1) >= maxLines.Value)
            {
                sb.AppendLine($"... (showed {maxLines.Value} of {totalSlides} slides, use --start/--end to see more)");
                break;
            }

            sb.AppendLine($"=== Slide {slideNum} ===");
            var shapes = GetSlide(slidePart).CommonSlideData?.ShapeTree?.Elements<Shape>() ?? Enumerable.Empty<Shape>();

            foreach (var shape in shapes)
            {
                var text = GetShapeText(shape);
                if (!string.IsNullOrWhiteSpace(text))
                    sb.AppendLine(text);
            }
            sb.AppendLine();
        }

        return sb.ToString().TrimEnd();
    }

    public string ViewAsAnnotated(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
    {
        var sb = new StringBuilder();
        int slideNum = 0;
        int totalSlides = GetSlideParts().Count();

        foreach (var slidePart in GetSlideParts())
        {
            slideNum++;
            if (startLine.HasValue && slideNum < startLine.Value) continue;
            if (endLine.HasValue && slideNum > endLine.Value) break;

            if (maxLines.HasValue && slideNum - (startLine ?? 1) >= maxLines.Value)
            {
                sb.AppendLine($"... (showed {maxLines.Value} of {totalSlides} slides, use --start/--end to see more)");
                break;
            }

            sb.AppendLine($"[Slide {slideNum}]");
            var shapes = GetSlide(slidePart).CommonSlideData?.ShapeTree?.ChildElements ?? Enumerable.Empty<OpenXmlElement>();

            int shapeIdx = 0;
            foreach (var child in shapes)
            {
                if (child is Shape shape)
                {
                    // Check if shape contains equations
                    var mathElements = FindShapeMathElements(shape);
                    if (mathElements.Count > 0)
                    {
                        var latex = string.Concat(mathElements.Select(FormulaParser.ToLatex));
                        var text = GetShapeText(shape);
                        // Check for text runs NOT inside mc:Fallback
                        var hasOtherText = shape.TextBody?.Elements<Drawing.Paragraph>()
                            .SelectMany(p => p.Elements<Drawing.Run>())
                            .Any(r => !string.IsNullOrWhiteSpace(r.Text?.Text)) == true;
                        if (hasOtherText)
                            sb.AppendLine($"  [Text Box] \"{text}\" \u2190 contains equation: \"{latex}\"");
                        else
                            sb.AppendLine($"  [Equation] \"{latex}\"");
                    }
                    else
                    {
                        var text = GetShapeText(shape);
                        var type = IsTitle(shape) ? "Title" : "Text Box";

                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            var firstRun = shape.TextBody?.Descendants<Drawing.Run>().FirstOrDefault();
                            var font = firstRun?.RunProperties?.GetFirstChild<Drawing.LatinFont>()?.Typeface
                                ?? firstRun?.RunProperties?.GetFirstChild<Drawing.EastAsianFont>()?.Typeface
                                ?? "(default)";
                            var fontSize = firstRun?.RunProperties?.FontSize?.Value;
                            var sizeStr = fontSize.HasValue ? $"{fontSize.Value / 100}pt" : "";

                            sb.AppendLine($"  [{type}] \"{text}\" \u2190 {font} {sizeStr}");
                        }
                    }
                    shapeIdx++;
                }
                else if (child is Picture pic)
                {
                    var name = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Picture";
                    var altText = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Description?.Value;
                    var altInfo = string.IsNullOrEmpty(altText) ? "\u26a0 no alt text" : $"alt=\"{altText}\"";
                    sb.AppendLine($"  [Picture] \"{name}\" \u2190 {altInfo}");
                }
            }
            sb.AppendLine();
        }

        return sb.ToString().TrimEnd();
    }

    public string ViewAsOutline()
    {
        var sb = new StringBuilder();
        var slideParts = GetSlideParts().ToList();

        sb.AppendLine($"File: {Path.GetFileName(_filePath)} | {slideParts.Count} slides");

        int slideNum = 0;
        foreach (var slidePart in slideParts)
        {
            slideNum++;
            var shapes = GetSlide(slidePart).CommonSlideData?.ShapeTree?.Elements<Shape>() ?? Enumerable.Empty<Shape>();

            var title = shapes.Where(IsTitle).Select(GetShapeText).FirstOrDefault(t => !string.IsNullOrWhiteSpace(t)) ?? "(untitled)";

            int textBoxes = shapes.Count(s => !IsTitle(s) && !string.IsNullOrWhiteSpace(GetShapeText(s)));
            int pictures = GetSlide(slidePart).CommonSlideData?.ShapeTree?.Elements<Picture>().Count() ?? 0;

            var details = new List<string>();
            if (textBoxes > 0) details.Add($"{textBoxes} text box(es)");
            if (pictures > 0) details.Add($"{pictures} picture(s)");

            var detailStr = details.Count > 0 ? $" - {string.Join(", ", details)}" : "";
            sb.AppendLine($"\u251c\u2500\u2500 Slide {slideNum}: \"{title}\"{detailStr}");
        }

        return sb.ToString().TrimEnd();
    }

    public string ViewAsStats()
    {
        var sb = new StringBuilder();
        var slideParts = GetSlideParts().ToList();

        int totalShapes = 0;
        int totalPictures = 0;
        int totalTextBoxes = 0;
        int slidesWithoutTitle = 0;
        int picturesWithoutAlt = 0;
        var fontCounts = new Dictionary<string, int>();

        foreach (var slidePart in slideParts)
        {
            var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
            if (shapeTree == null) continue;

            var shapes = shapeTree.Elements<Shape>().ToList();
            var pictures = shapeTree.Elements<Picture>().ToList();
            totalShapes += shapes.Count;
            totalPictures += pictures.Count;
            totalTextBoxes += shapes.Count(s => !IsTitle(s));

            if (!shapes.Any(IsTitle))
                slidesWithoutTitle++;

            picturesWithoutAlt += pictures.Count(p =>
                string.IsNullOrEmpty(p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Description?.Value));

            // Collect font usage
            foreach (var shape in shapes)
            {
                foreach (var run in shape.Descendants<Drawing.Run>())
                {
                    var font = run.RunProperties?.GetFirstChild<Drawing.LatinFont>()?.Typeface
                        ?? run.RunProperties?.GetFirstChild<Drawing.EastAsianFont>()?.Typeface;
                    if (font != null)
                        fontCounts[font!] = fontCounts.GetValueOrDefault(font!) + 1;
                }
            }
        }

        sb.AppendLine($"Slides: {slideParts.Count}");
        sb.AppendLine($"Total shapes: {totalShapes}");
        sb.AppendLine($"Text boxes: {totalTextBoxes}");
        sb.AppendLine($"Pictures: {totalPictures}");
        sb.AppendLine($"Slides without title: {slidesWithoutTitle}");
        sb.AppendLine($"Pictures without alt text: {picturesWithoutAlt}");

        if (fontCounts.Count > 0)
        {
            sb.AppendLine();
            sb.AppendLine("Font usage:");
            foreach (var (font, count) in fontCounts.OrderByDescending(kv => kv.Value))
                sb.AppendLine($"  {font}: {count} occurrence(s)");
        }

        return sb.ToString().TrimEnd();
    }

    public List<DocumentIssue> ViewAsIssues(string? issueType = null, int? limit = null)
    {
        var issues = new List<DocumentIssue>();
        int issueNum = 0;
        int slideNum = 0;

        foreach (var slidePart in GetSlideParts())
        {
            slideNum++;
            var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
            if (shapeTree == null) continue;

            var shapes = shapeTree.Elements<Shape>().ToList();
            if (!shapes.Any(IsTitle))
            {
                issues.Add(new DocumentIssue
                {
                    Id = $"S{++issueNum}",
                    Type = IssueType.Structure,
                    Severity = IssueSeverity.Warning,
                    Path = $"/slide[{slideNum}]",
                    Message = "Slide has no title"
                });
            }

            // Check for font consistency issues
            int shapeIdx = 0;
            foreach (var shape in shapes)
            {
                var runs = shape.Descendants<Drawing.Run>().ToList();
                if (runs.Count <= 1) { shapeIdx++; continue; }

                var fonts = runs.Select(r =>
                    r.RunProperties?.GetFirstChild<Drawing.LatinFont>()?.Typeface
                    ?? r.RunProperties?.GetFirstChild<Drawing.EastAsianFont>()?.Typeface)
                    .Where(f => f != null).Distinct().ToList();

                if (fonts.Count > 1)
                {
                    issues.Add(new DocumentIssue
                    {
                        Id = $"F{++issueNum}",
                        Type = IssueType.Format,
                        Severity = IssueSeverity.Info,
                        Path = $"/slide[{slideNum}]/shape[{shapeIdx + 1}]",
                        Message = $"Inconsistent fonts in text box: {string.Join(", ", fonts)}"
                    });
                }
                shapeIdx++;
            }

            foreach (var pic in shapeTree.Elements<Picture>())
            {
                var alt = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Description?.Value;
                if (string.IsNullOrEmpty(alt))
                {
                    var name = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value ?? "?";
                    issues.Add(new DocumentIssue
                    {
                        Id = $"F{++issueNum}",
                        Type = IssueType.Format,
                        Severity = IssueSeverity.Info,
                        Path = $"/slide[{slideNum}]",
                        Message = $"Picture \"{name}\" is missing alt text (accessibility issue)"
                    });
                }
            }

            if (limit.HasValue && issues.Count >= limit.Value) break;
        }

        return issues;
    }

    // ==================== Query Layer ====================

    public DocumentNode Get(string path, int depth = 1)
    {
        if (path == "/" || path == "")
        {
            var node = new DocumentNode { Path = "/", Type = "presentation" };
            int slideNum = 0;
            foreach (var slidePart in GetSlideParts())
            {
                slideNum++;
                var title = GetSlide(slidePart).CommonSlideData?.ShapeTree?.Elements<Shape>()
                    .Where(IsTitle).Select(GetShapeText).FirstOrDefault() ?? "(untitled)";

                var slideNode = new DocumentNode
                {
                    Path = $"/slide[{slideNum}]",
                    Type = "slide",
                    Preview = title
                };

                if (depth > 0)
                {
                    slideNode.Children = GetSlideChildNodes(slidePart, slideNum, depth - 1);
                    slideNode.ChildCount = slideNode.Children.Count;
                }
                else
                {
                    var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
                    slideNode.ChildCount = (shapeTree?.Elements<Shape>().Count() ?? 0)
                        + (shapeTree?.Elements<Picture>().Count() ?? 0);
                }

                node.Children.Add(slideNode);
            }
            node.ChildCount = node.Children.Count;
            return node;
        }

        // Parse /slide[N] or /slide[N]/shape[M]
        var match = Regex.Match(path, @"^/slide\[(\d+)\](?:/(\w+)\[(\d+)\])?$");
        if (!match.Success)
        {
            // Generic XML fallback: navigate by element localName
            var allSegments = GenericXmlQuery.ParsePathSegments(path);
            if (allSegments.Count == 0 || !allSegments[0].Name.Equals("slide", StringComparison.OrdinalIgnoreCase) || !allSegments[0].Index.HasValue)
                throw new ArgumentException($"Path must start with /slide[N]: {path}");

            var fbSlideIdx = allSegments[0].Index!.Value;
            var fbSlideParts = GetSlideParts().ToList();
            if (fbSlideIdx < 1 || fbSlideIdx > fbSlideParts.Count)
                throw new ArgumentException($"Slide {fbSlideIdx} not found (total: {fbSlideParts.Count})");

            OpenXmlElement fbCurrent = GetSlide(fbSlideParts[fbSlideIdx - 1]);
            var remaining = allSegments.Skip(1).ToList();
            if (remaining.Count > 0)
            {
                var target = GenericXmlQuery.NavigateByPath(fbCurrent, remaining);
                if (target == null)
                    return new DocumentNode { Path = path, Type = "error", Text = $"Element not found: {path}" };
                return GenericXmlQuery.ElementToNode(target, path, depth);
            }
            return GenericXmlQuery.ElementToNode(fbCurrent, path, depth);
        }

        var slideIdx = int.Parse(match.Groups[1].Value);
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");

        var targetSlidePart = slideParts[slideIdx - 1];

        if (!match.Groups[2].Success)
        {
            // Return slide node
            var slideNode = new DocumentNode
            {
                Path = path,
                Type = "slide",
                Preview = GetSlide(targetSlidePart).CommonSlideData?.ShapeTree?.Elements<Shape>()
                    .Where(IsTitle).Select(GetShapeText).FirstOrDefault() ?? "(untitled)"
            };
            slideNode.Children = GetSlideChildNodes(targetSlidePart, slideIdx, depth);
            slideNode.ChildCount = slideNode.Children.Count;
            return slideNode;
        }

        // Shape or picture
        var elementType = match.Groups[2].Value;
        var elementIdx = int.Parse(match.Groups[3].Value);
        var shapeTreeEl = GetSlide(targetSlidePart).CommonSlideData?.ShapeTree;
        if (shapeTreeEl == null)
            throw new ArgumentException($"Slide {slideIdx} has no shapes");

        if (elementType == "shape")
        {
            var shapes = shapeTreeEl.Elements<Shape>().ToList();
            if (elementIdx < 1 || elementIdx > shapes.Count)
                throw new ArgumentException($"Shape {elementIdx} not found (total: {shapes.Count})");
            return ShapeToNode(shapes[elementIdx - 1], slideIdx, elementIdx, depth);
        }
        else if (elementType == "picture" || elementType == "pic")
        {
            var pics = shapeTreeEl.Elements<Picture>().ToList();
            if (elementIdx < 1 || elementIdx > pics.Count)
                throw new ArgumentException($"Picture {elementIdx} not found (total: {pics.Count})");
            return PictureToNode(pics[elementIdx - 1], slideIdx, elementIdx);
        }

        // Generic fallback for unknown element types
        {
            var shapes2 = shapeTreeEl.ChildElements
                .Where(e => e.LocalName.Equals(elementType, StringComparison.OrdinalIgnoreCase)).ToList();
            if (elementIdx < 1 || elementIdx > shapes2.Count)
                throw new ArgumentException($"{elementType} {elementIdx} not found (total: {shapes2.Count})");
            return GenericXmlQuery.ElementToNode(shapes2[elementIdx - 1], path, depth);
        }
    }

    public List<DocumentNode> Query(string selector)
    {
        var results = new List<DocumentNode>();
        var parsed = ParseShapeSelector(selector);
        bool isEquationSelector = parsed.ElementType is "equation" or "math" or "formula";

        // Scheme B: generic XML fallback for unrecognized element types
        // Check if selector has a type that ParseShapeSelector didn't recognize
        var typeMatch = Regex.Match(selector.Contains(']') ? selector.Split(']').Last() : selector, @"^(?:slide\[\d+\]\s*>?\s*)?([\w:]+)");
        var rawType = typeMatch.Success ? typeMatch.Groups[1].Value.ToLowerInvariant() : "";
        bool isKnownType = string.IsNullOrEmpty(rawType)
            || rawType is "shape" or "textbox" or "title" or "picture" or "pic"
                or "equation" or "math" or "formula";
        if (!isKnownType)
        {
            var genericParsed = GenericXmlQuery.ParseSelector(selector);
            foreach (var slidePart in GetSlideParts())
            {
                results.AddRange(GenericXmlQuery.Query(
                    GetSlide(slidePart), genericParsed.element, genericParsed.attrs, genericParsed.containsText));
            }
            return results;
        }

        int slideNum = 0;

        foreach (var slidePart in GetSlideParts())
        {
            slideNum++;

            // Slide filter
            if (parsed.SlideNum.HasValue && parsed.SlideNum.Value != slideNum)
                continue;

            var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
            if (shapeTree == null) continue;

            int shapeIdx = 0;
            foreach (var shape in shapeTree.Elements<Shape>())
            {
                if (isEquationSelector)
                {
                    var mathElements = FindShapeMathElements(shape);
                    foreach (var mathElem in mathElements)
                    {
                        var latex = FormulaParser.ToLatex(mathElem);
                        if (parsed.TextContains == null || latex.Contains(parsed.TextContains))
                        {
                            results.Add(new DocumentNode
                            {
                                Path = $"/slide[{slideNum}]/shape[{shapeIdx + 1}]",
                                Type = "equation",
                                Text = latex,
                                Format = { ["mode"] = "display" }
                            });
                        }
                    }
                }
                else if (MatchesShapeSelector(shape, parsed))
                {
                    results.Add(ShapeToNode(shape, slideNum, shapeIdx + 1, 0));
                }
                shapeIdx++;
            }

            if (parsed.ElementType == "picture" || parsed.ElementType == "pic" || parsed.ElementType == null)
            {
                int picIdx = 0;
                foreach (var pic in shapeTree.Elements<Picture>())
                {
                    if (MatchesPictureSelector(pic, parsed))
                    {
                        results.Add(PictureToNode(pic, slideNum, picIdx + 1));
                    }
                    picIdx++;
                }
            }
        }

        return results;
    }

    public List<string> Set(string path, Dictionary<string, string> properties)
    {
        var match = Regex.Match(path, @"^/slide\[(\d+)\]/shape\[(\d+)\]$");
        if (!match.Success)
        {
            // Generic XML fallback: navigate to element and set attributes
            var allSegments = GenericXmlQuery.ParsePathSegments(path);
            if (allSegments.Count == 0 || !allSegments[0].Name.Equals("slide", StringComparison.OrdinalIgnoreCase) || !allSegments[0].Index.HasValue)
                throw new ArgumentException($"Path must start with /slide[N]: {path}");

            var fbSlideIdx = allSegments[0].Index!.Value;
            var fbSlideParts = GetSlideParts().ToList();
            if (fbSlideIdx < 1 || fbSlideIdx > fbSlideParts.Count)
                throw new ArgumentException($"Slide {fbSlideIdx} not found");

            var fbSlidePart = fbSlideParts[fbSlideIdx - 1];
            var remaining = allSegments.Skip(1).ToList();
            OpenXmlElement target = GetSlide(fbSlidePart);
            if (remaining.Count > 0)
            {
                target = GenericXmlQuery.NavigateByPath(target, remaining)
                    ?? throw new ArgumentException($"Element not found: {path}");
            }

            var unsup = new List<string>();
            foreach (var (key, value) in properties)
            {
                if (!GenericXmlQuery.SetGenericAttribute(target, key, value))
                    unsup.Add(key);
            }
            GetSlide(fbSlidePart).Save();
            return unsup;
        }

        var slideIdx = int.Parse(match.Groups[1].Value);
        var shapeIdx = int.Parse(match.Groups[2].Value);

        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found");

        var slidePart = slideParts[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
        if (shapeTree == null)
            throw new ArgumentException($"Slide {slideIdx} has no shapes");

        var shapes = shapeTree.Elements<Shape>().ToList();
        if (shapeIdx < 1 || shapeIdx > shapes.Count)
            throw new ArgumentException($"Shape {shapeIdx} not found");

        var shape = shapes[shapeIdx - 1];
        var unsupported = new List<string>();

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "text":
                    // Replace all text in the shape
                    var textBody = shape.TextBody;
                    if (textBody != null)
                    {
                        // Preserve formatting of first run, replace text
                        var firstRun = textBody.Descendants<Drawing.Run>().FirstOrDefault();
                        var runProps = firstRun?.RunProperties?.CloneNode(true) as Drawing.RunProperties;

                        // Remove all paragraphs
                        textBody.RemoveAllChildren<Drawing.Paragraph>();

                        // Add new paragraph with text
                        var newPara = new Drawing.Paragraph();
                        var newRun = new Drawing.Run();
                        if (runProps != null)
                            newRun.RunProperties = runProps;
                        newRun.Text = new Drawing.Text(value);
                        newPara.Append(newRun);
                        textBody.Append(newPara);
                    }
                    break;

                case "font":
                    foreach (var run in shape.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        // Remove existing font elements
                        rProps.RemoveAllChildren<Drawing.LatinFont>();
                        rProps.RemoveAllChildren<Drawing.EastAsianFont>();
                        rProps.RemoveAllChildren<Drawing.ComplexScriptFont>();
                        // Add new font
                        rProps.Append(new Drawing.LatinFont { Typeface = value });
                        rProps.Append(new Drawing.EastAsianFont { Typeface = value });
                    }
                    break;

                case "size":
                    var sizeVal = int.Parse(value) * 100; // pt to hundredths of a point
                    foreach (var run in shape.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.FontSize = sizeVal;
                    }
                    break;

                case "bold":
                    var isBold = bool.Parse(value);
                    foreach (var run in shape.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Bold = isBold;
                    }
                    break;

                case "italic":
                    var isItalic = bool.Parse(value);
                    foreach (var run in shape.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Italic = isItalic;
                    }
                    break;

                case "color":
                    foreach (var run in shape.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.RemoveAllChildren<Drawing.SolidFill>();
                        var solidFill = new Drawing.SolidFill();
                        solidFill.Append(new Drawing.RgbColorModelHex { Val = value });
                        // Use schema-aware insertion for correct element ordering
                        if (rProps is OpenXmlCompositeElement composite)
                        {
                            if (!composite.AddChild(solidFill, throwOnError: false))
                                rProps.AppendChild(solidFill);
                        }
                        else
                        {
                            rProps.AppendChild(solidFill);
                        }
                    }
                    break;

                default:
                    if (!GenericXmlQuery.SetGenericAttribute(shape, key, value))
                        unsupported.Add(key);
                    break;
            }
        }

        GetSlide(slidePart).Save();
        return unsupported;
    }

    public string Add(string parentPath, string type, int? index, Dictionary<string, string> properties)
    {
        switch (type.ToLowerInvariant())
        {
            case "slide":
                var presentationPart = _doc.PresentationPart
                    ?? throw new InvalidOperationException("Presentation not found");
                var presentation = presentationPart.Presentation
                    ?? throw new InvalidOperationException("No presentation");
                var slideIdList = presentation.GetFirstChild<SlideIdList>()
                    ?? presentation.AppendChild(new SlideIdList());

                var newSlidePart = presentationPart.AddNewPart<SlidePart>();

                // Link slide to slideLayout (required by PowerPoint)
                var slideMasterPart = presentationPart.SlideMasterParts.FirstOrDefault();
                if (slideMasterPart != null)
                {
                    var slideLayoutPart = slideMasterPart.SlideLayoutParts.FirstOrDefault();
                    if (slideLayoutPart != null)
                    {
                        newSlidePart.AddPart(slideLayoutPart);
                    }
                }

                newSlidePart.Slide = new Slide(
                    new CommonSlideData(
                        new ShapeTree(
                            new NonVisualGroupShapeProperties(
                                new NonVisualDrawingProperties { Id = 1, Name = "" },
                                new NonVisualGroupShapeDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()),
                            new GroupShapeProperties()
                        )
                    )
                );

                // Add title shape if text provided
                if (properties.TryGetValue("title", out var titleText))
                {
                    var titleShape = CreateTextShape(1, "Title", titleText, true);
                    newSlidePart.Slide.CommonSlideData!.ShapeTree!.AppendChild(titleShape);
                }

                // Add content text if provided
                if (properties.TryGetValue("text", out var contentText))
                {
                    var textShape = CreateTextShape(2, "Content", contentText, false);
                    newSlidePart.Slide.CommonSlideData!.ShapeTree!.AppendChild(textShape);
                }

                newSlidePart.Slide.Save();

                var maxId = slideIdList.Elements<SlideId>().Any()
                    ? slideIdList.Elements<SlideId>().Max(s => s.Id?.Value ?? 255) + 1
                    : 256;
                var relId = presentationPart.GetIdOfPart(newSlidePart);

                if (index.HasValue && index.Value < slideIdList.Elements<SlideId>().Count())
                {
                    var refSlide = slideIdList.Elements<SlideId>().ElementAtOrDefault(index.Value);
                    if (refSlide != null)
                        slideIdList.InsertBefore(new SlideId { Id = maxId, RelationshipId = relId }, refSlide);
                    else
                        slideIdList.AppendChild(new SlideId { Id = maxId, RelationshipId = relId });
                }
                else
                {
                    slideIdList.AppendChild(new SlideId { Id = maxId, RelationshipId = relId });
                }

                presentation.Save();
                var slideCount = slideIdList.Elements<SlideId>().Count();
                return $"/slide[{slideCount}]";

            case "shape" or "textbox":
                var slideMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]$");
                if (!slideMatch.Success)
                    throw new ArgumentException($"Shapes must be added to a slide: /slide[N]");

                var slideIdx = int.Parse(slideMatch.Groups[1].Value);
                var slideParts = GetSlideParts().ToList();
                if (slideIdx < 1 || slideIdx > slideParts.Count)
                    throw new ArgumentException($"Slide {slideIdx} not found");

                var slidePart = slideParts[slideIdx - 1];
                var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
                    ?? throw new InvalidOperationException("Slide has no shape tree");

                var text = properties.GetValueOrDefault("text", "");
                var shapeName = properties.GetValueOrDefault("name", $"TextBox {shapeTree.Elements<Shape>().Count() + 1}");
                var shapeId = (uint)(shapeTree.Elements<Shape>().Count() + shapeTree.Elements<Picture>().Count() + 2);

                var newShape = CreateTextShape(shapeId, shapeName, text, false);

                if (properties.TryGetValue("font", out var font))
                {
                    foreach (var run in newShape.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Append(new Drawing.LatinFont { Typeface = font });
                        rProps.Append(new Drawing.EastAsianFont { Typeface = font });
                    }
                }
                if (properties.TryGetValue("size", out var sizeStr))
                {
                    var sizeVal = int.Parse(sizeStr) * 100;
                    foreach (var run in newShape.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.FontSize = sizeVal;
                    }
                }
                if (properties.TryGetValue("bold", out var boldStr))
                {
                    var isBold = bool.Parse(boldStr);
                    foreach (var run in newShape.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Bold = isBold;
                    }
                }
                if (properties.TryGetValue("italic", out var italicStr))
                {
                    var isItalic = bool.Parse(italicStr);
                    foreach (var run in newShape.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.Italic = isItalic;
                    }
                }
                if (properties.TryGetValue("color", out var colorVal))
                {
                    foreach (var run in newShape.Descendants<Drawing.Run>())
                    {
                        var rProps = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rProps.RemoveAllChildren<Drawing.SolidFill>();
                        var solidFill = new Drawing.SolidFill();
                        solidFill.Append(new Drawing.RgbColorModelHex { Val = colorVal });
                        if (rProps is OpenXmlCompositeElement composite)
                        {
                            if (!composite.AddChild(solidFill, throwOnError: false))
                                rProps.AppendChild(solidFill);
                        }
                        else
                        {
                            rProps.AppendChild(solidFill);
                        }
                    }
                }

                // Position and size (in EMU, 1cm = 360000 EMU; or parse as cm/in)
                {
                    long xEmu = 0, yEmu = 0;
                    long cxEmu = 9144000, cyEmu = 742950; // default: ~25.4cm x ~2.06cm
                    if (properties.TryGetValue("x", out var xStr)) xEmu = ParseEmu(xStr);
                    if (properties.TryGetValue("y", out var yStr)) yEmu = ParseEmu(yStr);
                    if (properties.TryGetValue("width", out var wStr)) cxEmu = ParseEmu(wStr);
                    if (properties.TryGetValue("height", out var hStr)) cyEmu = ParseEmu(hStr);

                    newShape.ShapeProperties!.Transform2D = new Drawing.Transform2D
                    {
                        Offset = new Drawing.Offset { X = xEmu, Y = yEmu },
                        Extents = new Drawing.Extents { Cx = cxEmu, Cy = cyEmu }
                    };
                    newShape.ShapeProperties.AppendChild(
                        new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = Drawing.ShapeTypeValues.Rectangle }
                    );
                }

                shapeTree.AppendChild(newShape);
                GetSlide(slidePart).Save();
                var shapeCount = shapeTree.Elements<Shape>().Count();
                return $"/slide[{slideIdx}]/shape[{shapeCount}]";

            case "picture" or "image" or "img":
            {
                if (!properties.TryGetValue("path", out var imgPath))
                    throw new ArgumentException("'path' property is required for picture type");
                if (!File.Exists(imgPath))
                    throw new FileNotFoundException($"Image file not found: {imgPath}");

                var imgSlideMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]$");
                if (!imgSlideMatch.Success)
                    throw new ArgumentException($"Pictures must be added to a slide: /slide[N]");

                var imgSlideIdx = int.Parse(imgSlideMatch.Groups[1].Value);
                var imgSlideParts = GetSlideParts().ToList();
                if (imgSlideIdx < 1 || imgSlideIdx > imgSlideParts.Count)
                    throw new ArgumentException($"Slide {imgSlideIdx} not found");

                var imgSlidePart = imgSlideParts[imgSlideIdx - 1];
                var imgShapeTree = GetSlide(imgSlidePart).CommonSlideData?.ShapeTree
                    ?? throw new InvalidOperationException("Slide has no shape tree");

                // Determine image type
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

                // Embed image into slide part
                var imagePart = imgSlidePart.AddImagePart(imgPartType);
                using (var imgStream = File.OpenRead(imgPath))
                    imagePart.FeedData(imgStream);
                var imgRelId = imgSlidePart.GetIdOfPart(imagePart);

                // Dimensions (default: 6in x 4in)
                long cxEmu = 5486400; // 6 inches in EMUs
                long cyEmu = 3657600; // 4 inches in EMUs
                if (properties.TryGetValue("width", out var widthStr))
                    cxEmu = ParseEmu(widthStr);
                if (properties.TryGetValue("height", out var heightStr))
                    cyEmu = ParseEmu(heightStr);

                // Position (default: centered on standard 10x7.5 inch slide)
                long xEmu = (9144000 - cxEmu) / 2;
                long yEmu = (6858000 - cyEmu) / 2;
                if (properties.TryGetValue("x", out var xStr))
                    xEmu = ParseEmu(xStr);
                if (properties.TryGetValue("y", out var yStr))
                    yEmu = ParseEmu(yStr);

                var imgShapeId = (uint)(imgShapeTree.Elements<Shape>().Count() + imgShapeTree.Elements<Picture>().Count() + 2);
                var imgName = properties.GetValueOrDefault("name", $"Picture {imgShapeId}");
                var altText = properties.GetValueOrDefault("alt", Path.GetFileName(imgPath));

                // Build Picture element following Open-XML-SDK conventions
                var picture = new Picture();

                picture.NonVisualPictureProperties = new NonVisualPictureProperties(
                    new NonVisualDrawingProperties { Id = imgShapeId, Name = imgName, Description = altText },
                    new NonVisualPictureDrawingProperties(
                        new Drawing.PictureLocks { NoChangeAspect = true }
                    ),
                    new ApplicationNonVisualDrawingProperties()
                );

                picture.BlipFill = new BlipFill();
                picture.BlipFill.Blip = new Drawing.Blip { Embed = imgRelId };
                picture.BlipFill.AppendChild(new Drawing.Stretch(new Drawing.FillRectangle()));

                picture.ShapeProperties = new ShapeProperties();
                picture.ShapeProperties.Transform2D = new Drawing.Transform2D();
                picture.ShapeProperties.Transform2D.Offset = new Drawing.Offset { X = xEmu, Y = yEmu };
                picture.ShapeProperties.Transform2D.Extents = new Drawing.Extents { Cx = cxEmu, Cy = cyEmu };
                picture.ShapeProperties.AppendChild(
                    new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = Drawing.ShapeTypeValues.Rectangle }
                );

                imgShapeTree.AppendChild(picture);
                GetSlide(imgSlidePart).Save();

                var picCount = imgShapeTree.Elements<Picture>().Count();
                return $"/slide[{imgSlideIdx}]/picture[{picCount}]";
            }

            case "equation" or "formula" or "math":
            {
                if (!properties.TryGetValue("formula", out var eqFormula))
                    throw new ArgumentException("'formula' property is required for equation type");

                var eqSlideMatch = Regex.Match(parentPath, @"^/slide\[(\d+)\]$");
                if (!eqSlideMatch.Success)
                    throw new ArgumentException($"Equations must be added to a slide: /slide[N]");

                var eqSlideIdx = int.Parse(eqSlideMatch.Groups[1].Value);
                var eqSlideParts = GetSlideParts().ToList();
                if (eqSlideIdx < 1 || eqSlideIdx > eqSlideParts.Count)
                    throw new ArgumentException($"Slide {eqSlideIdx} not found");

                var eqSlidePart = eqSlideParts[eqSlideIdx - 1];
                var eqShapeTree = GetSlide(eqSlidePart).CommonSlideData?.ShapeTree
                    ?? throw new InvalidOperationException("Slide has no shape tree");

                var eqShapeId = (uint)(eqShapeTree.Elements<Shape>().Count() + eqShapeTree.Elements<Picture>().Count() + 2);
                var eqShapeName = properties.GetValueOrDefault("name", $"Equation {eqShapeId}");

                // Parse formula to OMML
                var mathContent = FormulaParser.Parse(eqFormula);
                M.OfficeMath oMath;
                if (mathContent is M.OfficeMath directMath)
                    oMath = directMath;
                else
                    oMath = new M.OfficeMath(mathContent.CloneNode(true));

                // Build the a14:m wrapper element via raw XML
                // PPT equations are embedded as: a:p > a14:m > m:oMathPara > m:oMath
                var mathPara = new M.Paragraph(oMath);
                var a14mXml = $"<a14:m xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\">{mathPara.OuterXml}</a14:m>";

                // Create shape with equation paragraph
                var eqShape = new Shape();
                eqShape.NonVisualShapeProperties = new NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = eqShapeId, Name = eqShapeName },
                    new NonVisualShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()
                );
                eqShape.ShapeProperties = new ShapeProperties();

                // Create text body with math paragraph
                var bodyProps = new Drawing.BodyProperties();
                var listStyle = new Drawing.ListStyle();
                var drawingPara = new Drawing.Paragraph();

                // Build mc:AlternateContent > mc:Choice(Requires="a14") > a14:m > m:oMathPara
                var a14mElement = new OpenXmlUnknownElement("a14", "m", "http://schemas.microsoft.com/office/drawing/2010/main");
                a14mElement.AppendChild(mathPara.CloneNode(true));

                var choice = new AlternateContentChoice();
                choice.Requires = "a14";
                choice.AppendChild(a14mElement);

                // Fallback: readable text for older versions
                var fallback = new AlternateContentFallback();
                var fallbackRun = new Drawing.Run(
                    new Drawing.RunProperties { Language = "en-US" },
                    new Drawing.Text(FormulaParser.ToReadableText(mathPara))
                );
                fallback.AppendChild(fallbackRun);

                var altContent = new AlternateContent();
                altContent.AppendChild(choice);
                altContent.AppendChild(fallback);
                drawingPara.AppendChild(altContent);

                eqShape.TextBody = new TextBody(bodyProps, listStyle, drawingPara);
                eqShapeTree.AppendChild(eqShape);
                GetSlide(eqSlidePart).Save();

                var eqShapeCount = eqShapeTree.Elements<Shape>().Count();
                return $"/slide[{eqSlideIdx}]/shape[{eqShapeCount}]";
            }

            default:
            {
                // Generic fallback: create typed element via SDK schema validation
                var allSegments = GenericXmlQuery.ParsePathSegments(parentPath);
                if (allSegments.Count == 0 || !allSegments[0].Name.Equals("slide", StringComparison.OrdinalIgnoreCase) || !allSegments[0].Index.HasValue)
                    throw new ArgumentException($"Generic add requires a path starting with /slide[N]: {parentPath}");

                var fbSlideIdx = allSegments[0].Index!.Value;
                var fbSlideParts = GetSlideParts().ToList();
                if (fbSlideIdx < 1 || fbSlideIdx > fbSlideParts.Count)
                    throw new ArgumentException($"Slide {fbSlideIdx} not found");

                var fbSlidePart = fbSlideParts[fbSlideIdx - 1];
                OpenXmlElement fbParent = GetSlide(fbSlidePart);
                var remaining = allSegments.Skip(1).ToList();
                if (remaining.Count > 0)
                {
                    fbParent = GenericXmlQuery.NavigateByPath(fbParent, remaining)
                        ?? throw new ArgumentException($"Parent element not found: {parentPath}");
                }

                var created = GenericXmlQuery.TryCreateTypedElement(fbParent, type, properties, index);
                if (created == null)
                    throw new ArgumentException($"Schema-invalid element type '{type}' for parent '{parentPath}'. " +
                        "Use raw-set --action append with explicit XML instead.");

                GetSlide(fbSlidePart).Save();

                // Build result path
                var siblings = fbParent.ChildElements.Where(e => e.LocalName == created.LocalName).ToList();
                var createdIdx = siblings.IndexOf(created) + 1;
                return $"{parentPath}/{created.LocalName}[{createdIdx}]";
            }
        }
    }

    public void Remove(string path)
    {
        var slideMatch = Regex.Match(path, @"^/slide\[(\d+)\](?:/(\w+)\[(\d+)\])?$");
        if (!slideMatch.Success)
            throw new ArgumentException($"Invalid path: {path}");

        var slideIdx = int.Parse(slideMatch.Groups[1].Value);

        if (!slideMatch.Groups[2].Success)
        {
            // Remove entire slide
            var presentationPart = _doc.PresentationPart
                ?? throw new InvalidOperationException("Presentation not found");
            var presentation = presentationPart.Presentation
                ?? throw new InvalidOperationException("No presentation");
            var slideIdList = presentation.GetFirstChild<SlideIdList>()
                ?? throw new InvalidOperationException("No slides");

            var slideIds = slideIdList.Elements<SlideId>().ToList();
            if (slideIdx < 1 || slideIdx > slideIds.Count)
                throw new ArgumentException($"Slide {slideIdx} not found");

            var slideId = slideIds[slideIdx - 1];
            var relId = slideId.RelationshipId?.Value;
            slideId.Remove();
            if (relId != null)
                presentationPart.DeletePart(presentationPart.GetPartById(relId));
            presentation.Save();
            return;
        }

        // Remove shape or picture from slide
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found");

        var slidePart = slideParts[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException("Slide has no shapes");

        var elementType = slideMatch.Groups[2].Value;
        var elementIdx = int.Parse(slideMatch.Groups[3].Value);

        if (elementType == "shape")
        {
            var shapes = shapeTree.Elements<Shape>().ToList();
            if (elementIdx < 1 || elementIdx > shapes.Count)
                throw new ArgumentException($"Shape {elementIdx} not found");
            shapes[elementIdx - 1].Remove();
        }
        else if (elementType is "picture" or "pic")
        {
            var pics = shapeTree.Elements<Picture>().ToList();
            if (elementIdx < 1 || elementIdx > pics.Count)
                throw new ArgumentException($"Picture {elementIdx} not found");
            pics[elementIdx - 1].Remove();
        }
        else
        {
            throw new ArgumentException($"Unknown element type: {elementType}");
        }

        GetSlide(slidePart).Save();
    }

    public string Move(string sourcePath, string? targetParentPath, int? index)
    {
        var presentationPart = _doc.PresentationPart
            ?? throw new InvalidOperationException("Presentation not found");
        var slideParts = GetSlideParts().ToList();

        // Case 1: Move entire slide (reorder)
        var slideOnlyMatch = Regex.Match(sourcePath, @"^/slide\[(\d+)\]$");
        if (slideOnlyMatch.Success)
        {
            var slideIdx = int.Parse(slideOnlyMatch.Groups[1].Value);
            var movePresentation = presentationPart.Presentation
                ?? throw new InvalidOperationException("No presentation");
            var slideIdList = movePresentation.GetFirstChild<SlideIdList>()
                ?? throw new InvalidOperationException("No slides");
            var slideIds = slideIdList.Elements<SlideId>().ToList();
            if (slideIdx < 1 || slideIdx > slideIds.Count)
                throw new ArgumentException($"Slide {slideIdx} not found");

            var slideId = slideIds[slideIdx - 1];
            slideId.Remove();

            if (index.HasValue)
            {
                var remaining = slideIdList.Elements<SlideId>().ToList();
                if (index.Value >= 0 && index.Value < remaining.Count)
                    remaining[index.Value].InsertBeforeSelf(slideId);
                else
                    slideIdList.AppendChild(slideId);
            }
            else
            {
                slideIdList.AppendChild(slideId);
            }

            movePresentation.Save();
            var newSlideIds = slideIdList.Elements<SlideId>().ToList();
            var newIdx = newSlideIds.IndexOf(slideId) + 1;
            return $"/slide[{newIdx}]";
        }

        // Case 2: Move element within/across slides
        var (srcSlidePart, srcElement) = ResolveSlideElement(sourcePath, slideParts);

        // Determine target
        string effectiveParentPath;
        SlidePart tgtSlidePart;
        ShapeTree tgtShapeTree;

        if (string.IsNullOrEmpty(targetParentPath))
        {
            // Reorder within same parent
            tgtSlidePart = srcSlidePart;
            tgtShapeTree = GetSlide(srcSlidePart).CommonSlideData?.ShapeTree
                ?? throw new InvalidOperationException("Slide has no shape tree");
            var srcSlideIdx = slideParts.IndexOf(srcSlidePart) + 1;
            effectiveParentPath = $"/slide[{srcSlideIdx}]";
        }
        else
        {
            effectiveParentPath = targetParentPath;
            var tgtSlideMatch = Regex.Match(targetParentPath, @"^/slide\[(\d+)\]$");
            if (!tgtSlideMatch.Success)
                throw new ArgumentException($"Target must be a slide: /slide[N]");
            var tgtSlideIdx = int.Parse(tgtSlideMatch.Groups[1].Value);
            if (tgtSlideIdx < 1 || tgtSlideIdx > slideParts.Count)
                throw new ArgumentException($"Slide {tgtSlideIdx} not found");
            tgtSlidePart = slideParts[tgtSlideIdx - 1];
            tgtShapeTree = GetSlide(tgtSlidePart).CommonSlideData?.ShapeTree
                ?? throw new InvalidOperationException("Slide has no shape tree");
        }

        srcElement.Remove();

        // Copy relationships if moving across slides
        if (srcSlidePart != tgtSlidePart)
            CopyRelationships(srcElement, srcSlidePart, tgtSlidePart);

        InsertAtPosition(tgtShapeTree, srcElement, index);

        GetSlide(srcSlidePart).Save();
        if (srcSlidePart != tgtSlidePart)
            GetSlide(tgtSlidePart).Save();

        return ComputeElementPath(effectiveParentPath, srcElement, tgtShapeTree);
    }

    public string CopyFrom(string sourcePath, string targetParentPath, int? index)
    {
        var slideParts = GetSlideParts().ToList();

        var (srcSlidePart, srcElement) = ResolveSlideElement(sourcePath, slideParts);
        var clone = srcElement.CloneNode(true);

        var tgtSlideMatch = Regex.Match(targetParentPath, @"^/slide\[(\d+)\]$");
        if (!tgtSlideMatch.Success)
            throw new ArgumentException($"Target must be a slide: /slide[N]");
        var tgtSlideIdx = int.Parse(tgtSlideMatch.Groups[1].Value);
        if (tgtSlideIdx < 1 || tgtSlideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {tgtSlideIdx} not found");

        var tgtSlidePart = slideParts[tgtSlideIdx - 1];
        var tgtShapeTree = GetSlide(tgtSlidePart).CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException("Slide has no shape tree");

        // Copy relationships if across slides
        if (srcSlidePart != tgtSlidePart)
            CopyRelationships(clone, srcSlidePart, tgtSlidePart);

        InsertAtPosition(tgtShapeTree, clone, index);
        GetSlide(tgtSlidePart).Save();

        return ComputeElementPath(targetParentPath, clone, tgtShapeTree);
    }

    private (SlidePart slidePart, OpenXmlElement element) ResolveSlideElement(string path, List<SlidePart> slideParts)
    {
        var match = Regex.Match(path, @"^/slide\[(\d+)\]/(\w+)\[(\d+)\]$");
        if (!match.Success)
            throw new ArgumentException($"Invalid element path: {path}. Expected /slide[N]/element[M]");

        var slideIdx = int.Parse(match.Groups[1].Value);
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"Slide {slideIdx} not found");

        var slidePart = slideParts[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException("Slide has no shape tree");

        var elementType = match.Groups[2].Value;
        var elementIdx = int.Parse(match.Groups[3].Value);

        OpenXmlElement element = elementType switch
        {
            "shape" => shapeTree.Elements<Shape>().ElementAtOrDefault(elementIdx - 1)
                ?? throw new ArgumentException($"Shape {elementIdx} not found"),
            "picture" or "pic" => shapeTree.Elements<Picture>().ElementAtOrDefault(elementIdx - 1)
                ?? throw new ArgumentException($"Picture {elementIdx} not found"),
            _ => shapeTree.ChildElements
                .Where(e => e.LocalName.Equals(elementType, StringComparison.OrdinalIgnoreCase))
                .ElementAtOrDefault(elementIdx - 1)
                ?? throw new ArgumentException($"{elementType} {elementIdx} not found")
        };

        return (slidePart, element);
    }

    private static void CopyRelationships(OpenXmlElement element, SlidePart sourcePart, SlidePart targetPart)
    {
        var rNsUri = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var allElements = element.Descendants().Prepend(element);

        foreach (var el in allElements.ToList())
        {
            foreach (var attr in el.GetAttributes().ToList())
            {
                if (attr.NamespaceUri != rNsUri) continue;

                var oldRelId = attr.Value;
                if (string.IsNullOrEmpty(oldRelId)) continue;

                try
                {
                    var referencedPart = sourcePart.GetPartById(oldRelId);
                    string newRelId;
                    try
                    {
                        newRelId = targetPart.GetIdOfPart(referencedPart);
                    }
                    catch
                    {
                        newRelId = targetPart.CreateRelationshipToPart(referencedPart);
                    }

                    if (newRelId != oldRelId)
                    {
                        el.SetAttribute(new OpenXmlAttribute(attr.Prefix, attr.LocalName, attr.NamespaceUri, newRelId));
                    }
                }
                catch { /* Not a valid relationship, skip */ }
            }
        }
    }

    private static void InsertAtPosition(OpenXmlElement parent, OpenXmlElement element, int? index)
    {
        if (index.HasValue && parent is ShapeTree)
        {
            // Skip structural elements (nvGrpSpPr, grpSpPr) that must stay at the beginning
            var contentChildren = parent.ChildElements
                .Where(e => e is not NonVisualGroupShapeProperties && e is not GroupShapeProperties)
                .ToList();
            if (index.Value >= 0 && index.Value < contentChildren.Count)
                contentChildren[index.Value].InsertBeforeSelf(element);
            else if (contentChildren.Count > 0)
                contentChildren.Last().InsertAfterSelf(element);
            else
                parent.AppendChild(element);
        }
        else if (index.HasValue)
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

    private static string ComputeElementPath(string parentPath, OpenXmlElement element, ShapeTree shapeTree)
    {
        // Map back to semantic type names
        string typeName;
        int typeIdx;
        if (element is Shape)
        {
            typeName = "shape";
            typeIdx = shapeTree.Elements<Shape>().ToList().IndexOf((Shape)element) + 1;
        }
        else if (element is Picture)
        {
            typeName = "picture";
            typeIdx = shapeTree.Elements<Picture>().ToList().IndexOf((Picture)element) + 1;
        }
        else
        {
            typeName = element.LocalName;
            typeIdx = shapeTree.ChildElements
                .Where(e => e.LocalName == element.LocalName)
                .ToList().IndexOf(element) + 1;
        }
        return $"{parentPath}/{typeName}[{typeIdx}]";
    }

    private static Shape CreateTextShape(uint id, string name, string text, bool isTitle)
    {
        var shape = new Shape();
        shape.NonVisualShapeProperties = new NonVisualShapeProperties(
            new NonVisualDrawingProperties { Id = id, Name = name },
            new NonVisualShapeDrawingProperties(),
            new ApplicationNonVisualDrawingProperties(
                isTitle ? new PlaceholderShape { Type = PlaceholderValues.Title } : new PlaceholderShape()
            )
        );
        shape.ShapeProperties = new ShapeProperties();
        shape.TextBody = new TextBody(
            new Drawing.BodyProperties(),
            new Drawing.ListStyle(),
            new Drawing.Paragraph(
                new Drawing.Run(
                    new Drawing.RunProperties { Language = "zh-CN" },
                    new Drawing.Text(text)
                )
            )
        );
        return shape;
    }

    // ==================== Raw Layer ====================

    public string Raw(string partPath, int? startRow = null, int? endRow = null, HashSet<string>? cols = null)
    {
        if (partPath == "/" || partPath == "/presentation")
            return _doc.PresentationPart?.Presentation?.OuterXml ?? "(empty)";

        var match = Regex.Match(partPath, @"^/slide\[(\d+)\]$");
        if (match.Success)
        {
            var idx = int.Parse(match.Groups[1].Value);
            var slideParts = GetSlideParts().ToList();
            if (idx >= 1 && idx <= slideParts.Count)
                return GetSlide(slideParts[idx - 1]).OuterXml;
            return $"(slide[{idx}] not found)";
        }

        return $"Unknown part: {partPath}. Available: /presentation, /slide[N]";
    }

    public void RawSet(string partPath, string xpath, string action, string? xml)
    {
        var presentationPart = _doc.PresentationPart
            ?? throw new InvalidOperationException("No presentation part");

        OpenXmlPartRootElement rootElement;

        if (partPath is "/" or "/presentation")
        {
            rootElement = presentationPart.Presentation
                ?? throw new InvalidOperationException("No presentation");
        }
        else if (Regex.Match(partPath, @"^/slide\[(\d+)\]$") is { Success: true } slideMatch)
        {
            var idx = int.Parse(slideMatch.Groups[1].Value);
            var slideParts = GetSlideParts().ToList();
            if (idx < 1 || idx > slideParts.Count)
                throw new ArgumentException($"Slide {idx} not found");
            rootElement = GetSlide(slideParts[idx - 1]);
        }
        else if (Regex.Match(partPath, @"^/slideMaster\[(\d+)\]$") is { Success: true } masterMatch)
        {
            var idx = int.Parse(masterMatch.Groups[1].Value);
            var masters = presentationPart.SlideMasterParts.ToList();
            if (idx < 1 || idx > masters.Count)
                throw new ArgumentException($"SlideMaster {idx} not found");
            rootElement = masters[idx - 1].SlideMaster
                ?? throw new InvalidOperationException("Corrupt file: slide master data missing");
        }
        else if (Regex.Match(partPath, @"^/slideLayout\[(\d+)\]$") is { Success: true } layoutMatch)
        {
            var idx = int.Parse(layoutMatch.Groups[1].Value);
            var layouts = presentationPart.SlideMasterParts
                .SelectMany(m => m.SlideLayoutParts).ToList();
            if (idx < 1 || idx > layouts.Count)
                throw new ArgumentException($"SlideLayout {idx} not found");
            rootElement = layouts[idx - 1].SlideLayout
                ?? throw new InvalidOperationException("Corrupt file: slide layout data missing");
        }
        else if (Regex.Match(partPath, @"^/noteSlide\[(\d+)\]$") is { Success: true } noteMatch)
        {
            var idx = int.Parse(noteMatch.Groups[1].Value);
            var slideParts = GetSlideParts().ToList();
            if (idx < 1 || idx > slideParts.Count)
                throw new ArgumentException($"Slide {idx} not found");
            var notesPart = slideParts[idx - 1].NotesSlidePart
                ?? throw new ArgumentException($"Slide {idx} has no notes");
            rootElement = notesPart.NotesSlide
                ?? throw new InvalidOperationException("Corrupt file: notes slide data missing");
        }
        else
        {
            throw new ArgumentException($"Unknown part: {partPath}. Available: /presentation, /slide[N], /slideMaster[N], /slideLayout[N], /noteSlide[N]");
        }

        var affected = RawXmlHelper.Execute(rootElement, xpath, action, xml);
        rootElement.Save();
        Console.WriteLine($"raw-set: {affected} element(s) affected");
    }

    public (string RelId, string PartPath) AddPart(string parentPartPath, string partType, Dictionary<string, string>? properties = null)
    {
        var presentationPart = _doc.PresentationPart
            ?? throw new InvalidOperationException("No presentation part");

        switch (partType.ToLowerInvariant())
        {
            case "chart":
                // Charts go under a SlidePart
                var slideMatch = System.Text.RegularExpressions.Regex.Match(
                    parentPartPath, @"^/slide\[(\d+)\]$");
                if (!slideMatch.Success)
                    throw new ArgumentException(
                        "Chart must be added under a slide: add-part <file> '/slide[N]' --type chart");

                var slideIdx = int.Parse(slideMatch.Groups[1].Value);
                var slideParts = GetSlideParts().ToList();
                if (slideIdx < 1 || slideIdx > slideParts.Count)
                    throw new ArgumentException($"Slide index {slideIdx} out of range");

                var slidePart = slideParts[slideIdx - 1];
                var chartPart = slidePart.AddNewPart<DocumentFormat.OpenXml.Packaging.ChartPart>();
                var relId = slidePart.GetIdOfPart(chartPart);

                chartPart.ChartSpace = new DocumentFormat.OpenXml.Drawing.Charts.ChartSpace(
                    new DocumentFormat.OpenXml.Drawing.Charts.Chart(
                        new DocumentFormat.OpenXml.Drawing.Charts.PlotArea(
                            new DocumentFormat.OpenXml.Drawing.Charts.Layout()
                        )
                    )
                );
                chartPart.ChartSpace.Save();

                var chartIdx = slidePart.ChartParts.ToList().IndexOf(chartPart);
                return (relId, $"/slide[{slideIdx}]/chart[{chartIdx + 1}]");

            default:
                throw new ArgumentException(
                    $"Unknown part type: {partType}. Supported: chart");
        }
    }

    public List<ValidationError> Validate() => RawXmlHelper.ValidateDocument(_doc);

    public void Dispose() => _doc.Dispose();

    // ==================== Private Helpers ====================

    private static Slide GetSlide(SlidePart part) =>
        part.Slide ?? throw new InvalidOperationException("Corrupt file: slide data missing");

    private IEnumerable<SlidePart> GetSlideParts()
    {
        var presentation = _doc.PresentationPart?.Presentation;
        var slideIdList = presentation?.GetFirstChild<SlideIdList>();
        if (slideIdList == null) yield break;

        foreach (var slideId in slideIdList.Elements<SlideId>())
        {
            var relId = slideId.RelationshipId?.Value;
            if (relId == null) continue;
            yield return (SlidePart)_doc.PresentationPart!.GetPartById(relId);
        }
    }

    private static string GetShapeText(Shape shape)
    {
        var textBody = shape.TextBody;
        if (textBody == null) return "";

        var sb = new StringBuilder();
        foreach (var para in textBody.Elements<Drawing.Paragraph>())
        {
            foreach (var child in para.ChildElements)
            {
                if (child is Drawing.Run run)
                    sb.Append(run.Text?.Text ?? "");
                else if (HasMathContent(child))
                    sb.Append(FormulaParser.ToReadableText(GetMathElement(child)));
            }
        }
        return sb.ToString();
    }

    /// <summary>
    /// Find all OMML math elements inside a shape's text body.
    /// </summary>
    private static List<OpenXmlElement> FindShapeMathElements(Shape shape)
    {
        var results = new List<OpenXmlElement>();
        var textBody = shape.TextBody;
        if (textBody == null) return results;

        foreach (var para in textBody.Elements<Drawing.Paragraph>())
        {
            foreach (var child in para.ChildElements)
            {
                if (HasMathContent(child))
                    results.Add(GetMathElement(child));
            }
        }
        return results;
    }

    /// <summary>
    /// Check if an element contains math content (a14:m or mc:AlternateContent with math).
    /// </summary>
    private static bool HasMathContent(OpenXmlElement element)
    {
        // Direct a14:m element
        if (element.LocalName == "m" && element.NamespaceUri == "http://schemas.microsoft.com/office/drawing/2010/main")
            return true;
        // mc:AlternateContent containing math (check both by type and localName)
        if (element is AlternateContent || element.LocalName == "AlternateContent")
        {
            // Check descendants for math, or check InnerXml
            if (element.Descendants().Any(e => e.LocalName == "oMath" || e.LocalName == "oMathPara"))
                return true;
            // Fallback: check raw XML for math namespace
            var innerXml = element.InnerXml;
            return innerXml.Contains("oMath");
        }
        return false;
    }

    /// <summary>
    /// Extract the OMML math element from an a14:m or mc:AlternateContent wrapper.
    /// </summary>
    private static OpenXmlElement GetMathElement(OpenXmlElement element)
    {
        // Direct a14:m → find oMath/oMathPara inside
        if (element.LocalName == "m" && element.NamespaceUri == "http://schemas.microsoft.com/office/drawing/2010/main")
        {
            // Try child elements first (works when element tree is properly parsed)
            var child = element.ChildElements.FirstOrDefault(e => e.LocalName == "oMathPara" || e.LocalName == "oMath");
            if (child != null) return child;

            // Try descendants
            var desc = element.Descendants().FirstOrDefault(e => e.LocalName == "oMathPara" || e.LocalName == "oMath");
            if (desc != null) return desc;

            // Last resort: re-parse from InnerXml (handles case where InnerXml was set but not parsed into children)
            var innerXml = element.InnerXml;
            if (!string.IsNullOrEmpty(innerXml) && innerXml.Contains("oMath"))
                return ReparseFromXml(innerXml) ?? element;

            return element;
        }
        // mc:AlternateContent → find oMath inside Choice
        if (element is AlternateContent || element.LocalName == "AlternateContent")
        {
            // Find Choice element (by type or localName)
            var choice = element.ChildElements.FirstOrDefault(e => e is AlternateContentChoice || e.LocalName == "Choice");
            if (choice != null)
            {
                var a14m = choice.ChildElements.FirstOrDefault(e =>
                    e.LocalName == "m" && e.NamespaceUri == "http://schemas.microsoft.com/office/drawing/2010/main");
                if (a14m != null)
                    return GetMathElement(a14m);

                // Try descendants directly
                var mathDesc = choice.Descendants().FirstOrDefault(e => e.LocalName == "oMathPara" || e.LocalName == "oMath");
                if (mathDesc != null)
                    return mathDesc;
            }

            // Fallback: try InnerXml parsing
            var innerXml = element.InnerXml;
            if (!string.IsNullOrEmpty(innerXml) && innerXml.Contains("oMath"))
                return ReparseFromXml(innerXml) ?? element;
        }
        return element;
    }

    /// <summary>
    /// Re-parse OMML XML string into an OpenXmlElement with navigable children.
    /// Uses OpenXmlUnknownElement which parses InnerXml into a proper child tree.
    /// </summary>
    private static OpenXmlElement? ReparseFromXml(string innerXml)
    {
        try
        {
            var xml = innerXml.Trim();
            // Find the outermost math element
            if (xml.Contains("oMathPara"))
            {
                // Extract the oMathPara element
                var startIdx = xml.IndexOf("<m:oMathPara", StringComparison.Ordinal);
                if (startIdx < 0) startIdx = xml.IndexOf("<oMathPara", StringComparison.Ordinal);
                if (startIdx >= 0)
                {
                    var endTag = xml.Contains("</m:oMathPara>") ? "</m:oMathPara>" : "</oMathPara>";
                    var endIdx = xml.IndexOf(endTag, StringComparison.Ordinal);
                    if (endIdx >= 0)
                    {
                        var oMathParaXml = xml[startIdx..(endIdx + endTag.Length)];
                        if (!oMathParaXml.Contains("xmlns:m="))
                            oMathParaXml = oMathParaXml.Replace("<m:oMathPara", "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"");
                        var wrapper = new OpenXmlUnknownElement("m", "oMathPara", "http://schemas.openxmlformats.org/officeDocument/2006/math");
                        // Extract inner content of oMathPara
                        var innerStart = oMathParaXml.IndexOf('>') + 1;
                        var innerEnd = oMathParaXml.LastIndexOf('<');
                        if (innerStart > 0 && innerEnd > innerStart)
                            wrapper.InnerXml = oMathParaXml[innerStart..innerEnd];
                        return wrapper;
                    }
                }
            }
        }
        catch
        {
            // Ignore parse failures
        }
        return null;
    }

    private static bool IsTitle(Shape shape)
    {
        var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
            ?.GetFirstChild<PlaceholderShape>();
        if (ph == null) return false;

        var type = ph.Type?.Value;
        return type == PlaceholderValues.Title || type == PlaceholderValues.CenteredTitle;
    }

    private static string GetShapeName(Shape shape)
    {
        return shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "?";
    }

    // ==================== Node Builders ====================

    private List<DocumentNode> GetSlideChildNodes(SlidePart slidePart, int slideNum, int depth)
    {
        var children = new List<DocumentNode>();
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
        if (shapeTree == null) return children;

        int shapeIdx = 0;
        foreach (var shape in shapeTree.Elements<Shape>())
        {
            children.Add(ShapeToNode(shape, slideNum, shapeIdx + 1, depth));
            shapeIdx++;
        }

        int picIdx = 0;
        foreach (var pic in shapeTree.Elements<Picture>())
        {
            children.Add(PictureToNode(pic, slideNum, picIdx + 1));
            picIdx++;
        }

        return children;
    }

    private static DocumentNode ShapeToNode(Shape shape, int slideNum, int shapeIdx, int depth)
    {
        var text = GetShapeText(shape);
        var name = GetShapeName(shape);
        var isTitle = IsTitle(shape);

        var node = new DocumentNode
        {
            Path = $"/slide[{slideNum}]/shape[{shapeIdx}]",
            Type = isTitle ? "title" : "textbox",
            Text = text,
            Preview = string.IsNullOrEmpty(text) ? name : (text.Length > 50 ? text[..50] + "..." : text)
        };

        node.Format["name"] = name;
        if (isTitle) node.Format["isTitle"] = true;

        // Collect font info
        var firstRun = shape.TextBody?.Descendants<Drawing.Run>().FirstOrDefault();
        if (firstRun?.RunProperties != null)
        {
            var font = firstRun.RunProperties.GetFirstChild<Drawing.LatinFont>()?.Typeface
                ?? firstRun.RunProperties.GetFirstChild<Drawing.EastAsianFont>()?.Typeface;
            if (font != null) node.Format["font"] = font;

            var fontSize = firstRun.RunProperties.FontSize?.Value;
            if (fontSize.HasValue) node.Format["size"] = $"{fontSize.Value / 100}pt";

            if (firstRun.RunProperties.Bold?.Value == true) node.Format["bold"] = true;
            if (firstRun.RunProperties.Italic?.Value == true) node.Format["italic"] = true;
        }

        // Count runs regardless of depth
        if (shape.TextBody != null)
        {
            var allRuns = shape.TextBody.Elements<Drawing.Paragraph>()
                .SelectMany(p => p.Elements<Drawing.Run>()).ToList();
            node.ChildCount = allRuns.Count;

            // Include individual runs at depth > 0
            if (depth > 0)
            {
                int runIdx = 0;
                foreach (var run in allRuns)
                {
                    var runNode = new DocumentNode
                    {
                        Path = $"/slide[{slideNum}]/shape[{shapeIdx}]/run[{runIdx}]",
                        Type = "run",
                        Text = run.Text?.Text ?? ""
                    };

                    if (run.RunProperties != null)
                    {
                        var f = run.RunProperties.GetFirstChild<Drawing.LatinFont>()?.Typeface
                            ?? run.RunProperties.GetFirstChild<Drawing.EastAsianFont>()?.Typeface;
                        if (f != null) runNode.Format["font"] = f;
                        var fs = run.RunProperties.FontSize?.Value;
                        if (fs.HasValue) runNode.Format["size"] = $"{fs.Value / 100}pt";
                        if (run.RunProperties.Bold?.Value == true) runNode.Format["bold"] = true;
                    }

                    node.Children.Add(runNode);
                    runIdx++;
                }
            }
        }

        return node;
    }

    private static DocumentNode PictureToNode(Picture pic, int slideNum, int picIdx)
    {
        var name = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Picture";
        var alt = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Description?.Value;

        var node = new DocumentNode
        {
            Path = $"/slide[{slideNum}]/picture[{picIdx}]",
            Type = "picture",
            Preview = name
        };

        node.Format["name"] = name;
        if (!string.IsNullOrEmpty(alt)) node.Format["alt"] = alt;
        else node.Format["alt"] = "(missing)";

        return node;
    }

    // ==================== Selector ====================

    private record ShapeSelector(string? ElementType, int? SlideNum, string? TextContains,
        string? FontEquals, string? FontNotEquals, bool? IsTitle, bool? HasAlt);

    private static ShapeSelector ParseShapeSelector(string selector)
    {
        string? elementType = null;
        int? slideNum = null;
        string? textContains = null;
        string? fontEquals = null;
        string? fontNotEquals = null;
        bool? isTitle = null;
        bool? hasAlt = null;

        // Check for slide prefix
        var slideMatch = Regex.Match(selector, @"slide\[(\d+)\]\s*(.*)");
        if (slideMatch.Success)
        {
            slideNum = int.Parse(slideMatch.Groups[1].Value);
            selector = slideMatch.Groups[2].Value.TrimStart('>', ' ');
        }

        // Element type
        var typeMatch = Regex.Match(selector, @"^(\w+)");
        if (typeMatch.Success)
        {
            var t = typeMatch.Groups[1].Value.ToLowerInvariant();
            if (t is "shape" or "textbox" or "title" or "picture" or "pic" or "equation" or "math" or "formula")
                elementType = t;
        }

        // Attributes
        foreach (Match attrMatch in Regex.Matches(selector, @"\[(\w+)(!?=)([^\]]*)\]"))
        {
            var key = attrMatch.Groups[1].Value.ToLowerInvariant();
            var op = attrMatch.Groups[2].Value;
            var val = attrMatch.Groups[3].Value.Trim('\'', '"');

            switch (key)
            {
                case "font" when op == "=": fontEquals = val; break;
                case "font" when op == "!=": fontNotEquals = val; break;
                case "title": isTitle = val.ToLowerInvariant() != "false"; break;
                case "alt": hasAlt = !string.IsNullOrEmpty(val) && val.ToLowerInvariant() != "false"; break;
            }
        }

        // :contains()
        var containsMatch = Regex.Match(selector, @":contains\(['""]?(.+?)['""]?\)");
        if (containsMatch.Success) textContains = containsMatch.Groups[1].Value;

        // Element type shortcuts
        if (elementType == "title") isTitle = true;

        // :no-alt
        if (selector.Contains(":no-alt")) hasAlt = false;

        return new ShapeSelector(elementType, slideNum, textContains, fontEquals, fontNotEquals, isTitle, hasAlt);
    }

    private static bool MatchesShapeSelector(Shape shape, ShapeSelector selector)
    {
        // Element type filter
        if (selector.ElementType is "picture" or "pic")
            return false;

        // Title filter
        if (selector.IsTitle.HasValue)
        {
            if (selector.IsTitle.Value != IsTitle(shape)) return false;
        }

        // Text contains
        if (selector.TextContains != null)
        {
            var text = GetShapeText(shape);
            if (!text.Contains(selector.TextContains, StringComparison.OrdinalIgnoreCase))
                return false;
        }

        // Font filter
        var runs = shape.Descendants<Drawing.Run>().ToList();
        if (selector.FontEquals != null)
        {
            bool found = runs.Any(r =>
            {
                var font = r.RunProperties?.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value
                    ?? r.RunProperties?.GetFirstChild<Drawing.EastAsianFont>()?.Typeface?.Value;
                return font != null && string.Equals(font, selector.FontEquals, StringComparison.OrdinalIgnoreCase);
            });
            if (!found) return false;
        }

        if (selector.FontNotEquals != null)
        {
            bool hasWrongFont = runs.Any(r =>
            {
                var font = r.RunProperties?.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value
                    ?? r.RunProperties?.GetFirstChild<Drawing.EastAsianFont>()?.Typeface?.Value;
                return font != null && !string.Equals(font, selector.FontNotEquals, StringComparison.OrdinalIgnoreCase);
            });
            if (!hasWrongFont) return false;
        }

        return true;
    }

    private static bool MatchesPictureSelector(Picture pic, ShapeSelector selector)
    {
        // Only match if looking for pictures specifically or no type specified
        if (selector.ElementType != null && selector.ElementType != "picture" && selector.ElementType != "pic")
            return false;

        if (selector.IsTitle.HasValue) return false; // Pictures can't be titles

        // Alt text filter
        if (selector.HasAlt.HasValue)
        {
            var alt = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Description?.Value;
            bool hasAlt = !string.IsNullOrEmpty(alt);
            if (selector.HasAlt.Value != hasAlt) return false;
        }

        return true;
    }

    private static long ParseEmu(string value)
    {
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
}
