// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public class ExcelHandler : IDocumentHandler
{
    private readonly SpreadsheetDocument _doc;
    private readonly string _filePath;

    public ExcelHandler(string filePath, bool editable)
    {
        _filePath = filePath;
        try
        {
            _doc = SpreadsheetDocument.Open(filePath, editable);
            // Force early validation: access WorkbookPart to catch corrupt packages now
            _ = _doc.WorkbookPart?.Workbook;
        }
        catch (DocumentFormat.OpenXml.Packaging.OpenXmlPackageException ex)
        {
            throw new InvalidOperationException(
                $"Cannot open {Path.GetFileName(filePath)}: {ex.Message}", ex);
        }
    }

    public string ViewAsText(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
    {
        var sb = new StringBuilder();
        var sheets = GetWorksheets();
        int sheetIdx = 0;
        int emitted = 0;
        bool truncated = false;

        foreach (var (sheetName, worksheetPart) in sheets)
        {
            if (truncated) break;
            sb.AppendLine($"=== Sheet: {sheetName} ===");
            var sheetData = GetSheet(worksheetPart).GetFirstChild<SheetData>();
            if (sheetData == null) continue;

            int totalRows = sheetData.Elements<Row>().Count();
            int lineNum = 0;
            foreach (var row in sheetData.Elements<Row>())
            {
                lineNum++;
                if (startLine.HasValue && lineNum < startLine.Value) continue;
                if (endLine.HasValue && lineNum > endLine.Value) break;

                if (maxLines.HasValue && emitted >= maxLines.Value)
                {
                    sb.AppendLine($"... (showed {emitted} rows, {totalRows} total in sheet, use --start/--end to view more)");
                    truncated = true;
                    break;
                }

                var cellElements = row.Elements<Cell>();
                if (cols != null)
                    cellElements = cellElements.Where(c => cols.Contains(ParseCellReference(c.CellReference?.Value ?? "A1").Column));
                var cells = cellElements.Select(c => GetCellDisplayValue(c)).ToArray();
                sb.AppendLine($"[{lineNum}] {string.Join("\t", cells)}");
                emitted++;
            }

            sheetIdx++;
            if (sheetIdx < sheets.Count) sb.AppendLine();
        }

        return sb.ToString().TrimEnd();
    }

    public string ViewAsAnnotated(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
    {
        var sb = new StringBuilder();
        var sheets = GetWorksheets();
        int emitted = 0;
        bool truncated = false;

        foreach (var (sheetName, worksheetPart) in sheets)
        {
            if (truncated) break;
            sb.AppendLine($"=== Sheet: {sheetName} ===");
            var sheetData = GetSheet(worksheetPart).GetFirstChild<SheetData>();
            if (sheetData == null) continue;

            int totalRows = sheetData.Elements<Row>().Count();
            int lineNum = 0;
            foreach (var row in sheetData.Elements<Row>())
            {
                lineNum++;
                if (startLine.HasValue && lineNum < startLine.Value) continue;
                if (endLine.HasValue && lineNum > endLine.Value) break;

                if (maxLines.HasValue && emitted >= maxLines.Value)
                {
                    sb.AppendLine($"... (showed {emitted} rows, {totalRows} total in sheet, use --start/--end to view more)");
                    truncated = true;
                    break;
                }

                var cellElements = row.Elements<Cell>();
                if (cols != null)
                    cellElements = cellElements.Where(c => cols.Contains(ParseCellReference(c.CellReference?.Value ?? "A1").Column));

                foreach (var cell in cellElements)
                {
                    var cellRef = cell.CellReference?.Value ?? "?";
                    var value = GetCellDisplayValue(cell);
                    var formula = cell.CellFormula?.Text;
                    var type = cell.DataType?.Value.ToString() ?? "Number";

                    var annotation = formula != null ? $"={formula}" : type;
                    var warn = "";

                    if (string.IsNullOrEmpty(value) && formula == null)
                        warn = " \u26a0 empty";
                    else if (formula != null && (value == "#REF!" || value == "#VALUE!" || value == "#NAME?"))
                        warn = " \u26a0 formula error";

                    sb.AppendLine($"  {cellRef}: [{value}] \u2190 {annotation}{warn}");
                }
                emitted++;
            }
        }

        return sb.ToString().TrimEnd();
    }

    public string ViewAsOutline()
    {
        var sb = new StringBuilder();
        var workbook = _doc.WorkbookPart?.Workbook;
        if (workbook == null) return "(empty workbook)";

        var sheets = workbook.GetFirstChild<Sheets>();
        if (sheets == null) return "(no sheets)";

        sb.AppendLine($"File: {Path.GetFileName(_filePath)}");

        foreach (var sheet in sheets.Elements<Sheet>())
        {
            var name = sheet.Name?.Value ?? "?";
            var sheetId = sheet.Id?.Value;
            if (sheetId == null) continue;

            var worksheetPart = (WorksheetPart)_doc.WorkbookPart!.GetPartById(sheetId);
            var sheetData = GetSheet(worksheetPart).GetFirstChild<SheetData>();

            int rowCount = sheetData?.Elements<Row>().Count() ?? 0;
            int colCount = sheetData?.Elements<Row>().FirstOrDefault()?.Elements<Cell>().Count() ?? 0;

            int formulaCount = 0;
            if (sheetData != null)
            {
                formulaCount = sheetData.Descendants<CellFormula>().Count();
            }

            var formulaInfo = formulaCount > 0 ? $", {formulaCount} formula(s)" : "";
            sb.AppendLine($"\u251c\u2500\u2500 \"{name}\" ({rowCount} rows \u00d7 {colCount} cols{formulaInfo})");
        }

        return sb.ToString().TrimEnd();
    }

    public string ViewAsStats()
    {
        var sb = new StringBuilder();
        var sheets = GetWorksheets();
        int totalCells = 0;
        int emptyCells = 0;
        int formulaCells = 0;
        int errorCells = 0;
        var typeCounts = new Dictionary<string, int>();

        foreach (var (sheetName, worksheetPart) in sheets)
        {
            var sheetData = GetSheet(worksheetPart).GetFirstChild<SheetData>();
            if (sheetData == null) continue;

            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    totalCells++;
                    var value = GetCellDisplayValue(cell);
                    if (string.IsNullOrEmpty(value)) emptyCells++;
                    if (cell.CellFormula != null) formulaCells++;
                    if (value is "#REF!" or "#VALUE!" or "#NAME?" or "#DIV/0!") errorCells++;

                    var type = cell.DataType?.Value.ToString() ?? "Number";
                    typeCounts[type] = typeCounts.GetValueOrDefault(type) + 1;
                }
            }
        }

        sb.AppendLine($"Sheets: {sheets.Count}");
        sb.AppendLine($"Total Cells: {totalCells}");
        sb.AppendLine($"Empty Cells: {emptyCells}");
        sb.AppendLine($"Formula Cells: {formulaCells}");
        sb.AppendLine($"Error Cells: {errorCells}");
        sb.AppendLine();
        sb.AppendLine("Data Type Distribution:");
        foreach (var (type, count) in typeCounts.OrderByDescending(kv => kv.Value))
            sb.AppendLine($"  {type}: {count}");

        return sb.ToString().TrimEnd();
    }

    public List<DocumentIssue> ViewAsIssues(string? issueType = null, int? limit = null)
    {
        var issues = new List<DocumentIssue>();
        int issueNum = 0;

        var sheets = GetWorksheets();
        foreach (var (sheetName, worksheetPart) in sheets)
        {
            var sheetData = GetSheet(worksheetPart).GetFirstChild<SheetData>();
            if (sheetData == null) continue;

            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    var cellRef = cell.CellReference?.Value ?? "?";
                    var value = GetCellDisplayValue(cell);

                    if (cell.CellFormula != null && value is "#REF!" or "#VALUE!" or "#NAME?" or "#DIV/0!")
                    {
                        issues.Add(new DocumentIssue
                        {
                            Id = $"F{++issueNum}",
                            Type = IssueType.Content,
                            Severity = IssueSeverity.Error,
                            Path = $"{sheetName}!{cellRef}",
                            Message = $"Formula error: {value}",
                            Context = $"={cell.CellFormula.Text}"
                        });
                    }

                    if (limit.HasValue && issues.Count >= limit.Value) break;
                }
                if (limit.HasValue && issues.Count >= limit.Value) break;
            }
        }

        return issues;
    }

    // ==================== Query Layer ====================

    public DocumentNode Get(string path, int depth = 1)
    {
        if (path == "/" || path == "")
        {
            var node = new DocumentNode { Path = "/", Type = "workbook" };
            foreach (var (name, part) in GetWorksheets())
            {
                var sheetNode = new DocumentNode { Path = $"/{name}", Type = "sheet", Preview = name };
                var sheetData = GetSheet(part).GetFirstChild<SheetData>();
                sheetNode.ChildCount = sheetData?.Elements<Row>().Count() ?? 0;

                if (depth > 0 && sheetData != null)
                {
                    sheetNode.Children = GetSheetChildNodes(name, sheetData, depth);
                }

                node.Children.Add(sheetNode);
            }
            node.ChildCount = node.Children.Count;
            return node;
        }

        // Parse path: /SheetName or /SheetName/A1 or /SheetName/A1:D10
        var segments = path.TrimStart('/').Split('/', 2);
        var sheetNameFromPath = segments[0];
        var worksheet = FindWorksheet(sheetNameFromPath);
        if (worksheet == null)
            throw new ArgumentException($"Sheet not found: {sheetNameFromPath}");

        var data = GetSheet(worksheet).GetFirstChild<SheetData>();
        if (data == null)
            return new DocumentNode { Path = path, Type = "sheet", Preview = "(empty)" };

        if (segments.Length == 1)
        {
            // Return sheet overview
            var sheetNode = new DocumentNode
            {
                Path = path,
                Type = "sheet",
                Preview = sheetNameFromPath,
                ChildCount = data.Elements<Row>().Count()
            };
            if (depth > 0)
            {
                sheetNode.Children = GetSheetChildNodes(sheetNameFromPath, data, depth);
            }
            return sheetNode;
        }

        // Cell reference: A1 or range A1:D10
        var cellRef = segments[1];

        // Check if it's a cell reference or a generic XML path
        var firstPart = cellRef.Split('/')[0].Split('[')[0];
        bool isCellRef = System.Text.RegularExpressions.Regex.IsMatch(firstPart, @"^[A-Z]+\d+", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        if (!isCellRef)
        {
            // Generic XML fallback: navigate worksheet XML tree
            var xmlSegments = GenericXmlQuery.ParsePathSegments(cellRef);
            var target = GenericXmlQuery.NavigateByPath(GetSheet(worksheet), xmlSegments);
            if (target == null)
                return new DocumentNode { Path = path, Type = "error", Text = $"Element not found: {cellRef}" };
            return GenericXmlQuery.ElementToNode(target, path, depth);
        }

        if (cellRef.Contains(':'))
        {
            // Range
            return GetCellRange(sheetNameFromPath, data, cellRef, depth);
        }
        else
        {
            // Single cell
            var cell = FindCell(data, cellRef);
            if (cell == null)
                return new DocumentNode { Path = path, Type = "cell", Text = "(empty)", Preview = cellRef };
            return CellToNode(sheetNameFromPath, cell);
        }
    }

    public List<DocumentNode> Query(string selector)
    {
        var results = new List<DocumentNode>();

        // Check if element type is known (Scheme A) or should fall back to generic XML (Scheme B)
        var elementMatch = Regex.Match(selector.Split('!').Last(), @"^([\w:]+)");
        var elementName = elementMatch.Success ? elementMatch.Groups[1].Value : "";
        bool isKnownType = string.IsNullOrEmpty(elementName)
            || elementName is "cell" or "row" or "sheet"
            || (elementName.Length <= 3 && Regex.IsMatch(elementName, @"^[A-Z]+$", RegexOptions.IgnoreCase));
        if (!isKnownType)
        {
            // Scheme B: generic XML fallback
            var genericParsed = GenericXmlQuery.ParseSelector(selector);
            foreach (var (_, worksheetPart) in GetWorksheets())
            {
                results.AddRange(GenericXmlQuery.Query(
                    GetSheet(worksheetPart), genericParsed.element, genericParsed.attrs, genericParsed.containsText));
            }
            return results;
        }

        var parsed = ParseCellSelector(selector);

        foreach (var (sheetName, worksheetPart) in GetWorksheets())
        {
            // If selector specifies a sheet, skip non-matching sheets
            if (parsed.Sheet != null && !sheetName.Equals(parsed.Sheet, StringComparison.OrdinalIgnoreCase))
                continue;

            var sheetData = GetSheet(worksheetPart).GetFirstChild<SheetData>();
            if (sheetData == null) continue;

            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    if (MatchesCellSelector(cell, sheetName, parsed))
                    {
                        results.Add(CellToNode(sheetName, cell));
                    }
                }
            }
        }

        return results;
    }

    public List<string> Set(string path, Dictionary<string, string> properties)
    {
        // Parse path: /SheetName/A1
        var segments = path.TrimStart('/').Split('/', 2);
        if (segments.Length < 2)
            throw new ArgumentException($"Path must include sheet and cell reference: /SheetName/A1");

        var sheetName = segments[0];
        var cellRef = segments[1];

        var worksheet = FindWorksheet(sheetName);
        if (worksheet == null)
            throw new ArgumentException($"Sheet not found: {sheetName}");

        // Check if path is a cell reference or generic XML path
        var firstPart = cellRef.Split('/')[0].Split('[')[0];
        bool isCellRef = System.Text.RegularExpressions.Regex.IsMatch(firstPart, @"^[A-Z]+\d+", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (!isCellRef)
        {
            // Generic XML fallback: navigate to element and set attributes
            var xmlSegments = GenericXmlQuery.ParsePathSegments(cellRef);
            var target = GenericXmlQuery.NavigateByPath(GetSheet(worksheet), xmlSegments);
            if (target == null)
                throw new ArgumentException($"Element not found: {cellRef}");
            var unsup = new List<string>();
            foreach (var (key, value) in properties)
            {
                if (!GenericXmlQuery.SetGenericAttribute(target, key, value))
                    unsup.Add(key);
            }
            GetSheet(worksheet).Save();
            return unsup;
        }

        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>();
        if (sheetData == null)
        {
            sheetData = new SheetData();
            GetSheet(worksheet).Append(sheetData);
        }

        var cell = FindOrCreateCell(sheetData, cellRef);

        // Separate content props from style props
        var styleProps = new Dictionary<string, string>();
        var unsupported = new List<string>();

        foreach (var (key, value) in properties)
        {
            if (ExcelStyleManager.IsStyleKey(key))
            {
                styleProps[key] = value;
                continue;
            }

            switch (key.ToLowerInvariant())
            {
                case "value":
                    cell.CellValue = new CellValue(value);
                    // Auto-detect type
                    if (double.TryParse(value, out _))
                        cell.DataType = null; // Number is default
                    else
                    {
                        cell.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.String);
                    }
                    break;
                case "formula":
                    cell.CellFormula = new CellFormula(value);
                    cell.CellValue = null;
                    break;
                case "type":
                    cell.DataType = value.ToLowerInvariant() switch
                    {
                        "string" or "str" => new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.String),
                        "number" or "num" => null,
                        "boolean" or "bool" => new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.Boolean),
                        _ => cell.DataType
                    };
                    break;
                case "clear":
                    cell.CellValue = null;
                    cell.CellFormula = null;
                    break;
                default:
                    if (!GenericXmlQuery.SetGenericAttribute(cell, key, value))
                        unsupported.Add(key);
                    break;
            }
        }

        // Apply style properties if any
        if (styleProps.Count > 0)
        {
            var workbookPart = _doc.WorkbookPart
                ?? throw new InvalidOperationException("Workbook not found");
            var styleManager = new ExcelStyleManager(workbookPart);
            cell.StyleIndex = styleManager.ApplyStyle(cell, styleProps);
        }

        GetSheet(worksheet).Save();
        return unsupported;
    }

    public string Add(string parentPath, string type, int? index, Dictionary<string, string> properties)
    {
        switch (type.ToLowerInvariant())
        {
            case "sheet":
                var workbookPart = _doc.WorkbookPart
                    ?? throw new InvalidOperationException("Workbook not found");
                var sheets = GetWorkbook().GetFirstChild<Sheets>()
                    ?? GetWorkbook().AppendChild(new Sheets());

                var name = properties.GetValueOrDefault("name", $"Sheet{sheets.Elements<Sheet>().Count() + 1}");
                var newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                newWorksheetPart.Worksheet = new Worksheet(new SheetData());
                newWorksheetPart.Worksheet.Save();

                var sheetId = sheets.Elements<Sheet>().Any()
                    ? sheets.Elements<Sheet>().Max(s => s.SheetId?.Value ?? 0) + 1
                    : 1;
                var relId = workbookPart.GetIdOfPart(newWorksheetPart);

                sheets.AppendChild(new Sheet { Id = relId, SheetId = (uint)sheetId, Name = name });
                GetWorkbook().Save();
                return $"/{name}";

            case "row":
                var segments = parentPath.TrimStart('/').Split('/', 2);
                var sheetName = segments[0];
                var worksheet = FindWorksheet(sheetName)
                    ?? throw new ArgumentException($"Sheet not found: {sheetName}");
                var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>()
                    ?? GetSheet(worksheet).AppendChild(new SheetData());

                var rowIdx = index ?? ((int)(sheetData.Elements<Row>().LastOrDefault()?.RowIndex?.Value ?? 0) + 1);
                var newRow = new Row { RowIndex = (uint)rowIdx };

                // Create cells if cols specified
                if (properties.TryGetValue("cols", out var colsStr))
                {
                    var cols = int.Parse(colsStr);
                    for (int c = 0; c < cols; c++)
                    {
                        var colLetter = IndexToColumnName(c + 1);
                        newRow.AppendChild(new Cell { CellReference = $"{colLetter}{rowIdx}" });
                    }
                }

                var afterRow = sheetData.Elements<Row>().LastOrDefault(r => (r.RowIndex?.Value ?? 0) < (uint)rowIdx);
                if (afterRow != null)
                    afterRow.InsertAfterSelf(newRow);
                else
                    sheetData.InsertAt(newRow, 0);

                GetSheet(worksheet).Save();
                return $"/{sheetName}/row[{rowIdx}]";

            case "cell":
                var cellSegments = parentPath.TrimStart('/').Split('/', 2);
                var cellSheetName = cellSegments[0];
                var cellWorksheet = FindWorksheet(cellSheetName)
                    ?? throw new ArgumentException($"Sheet not found: {cellSheetName}");
                var cellSheetData = GetSheet(cellWorksheet).GetFirstChild<SheetData>()
                    ?? GetSheet(cellWorksheet).AppendChild(new SheetData());

                var cellRef = properties.GetValueOrDefault("ref", "A1");
                var cell = FindOrCreateCell(cellSheetData, cellRef);

                if (properties.TryGetValue("value", out var value))
                {
                    cell.CellValue = new CellValue(value);
                    if (!double.TryParse(value, out _))
                        cell.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.String);
                }
                if (properties.TryGetValue("formula", out var formula))
                {
                    cell.CellFormula = new CellFormula(formula);
                    cell.CellValue = null;
                }
                if (properties.TryGetValue("type", out var cellType))
                {
                    cell.DataType = cellType.ToLowerInvariant() switch
                    {
                        "string" or "str" => new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.String),
                        "number" or "num" => null,
                        "boolean" or "bool" => new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.Boolean),
                        _ => cell.DataType
                    };
                }
                if (properties.TryGetValue("clear", out _))
                {
                    cell.CellValue = null;
                    cell.CellFormula = null;
                }

                // Apply style properties if any
                var cellStyleProps = new Dictionary<string, string>();
                foreach (var (key, val) in properties)
                {
                    if (ExcelStyleManager.IsStyleKey(key))
                        cellStyleProps[key] = val;
                }
                if (cellStyleProps.Count > 0)
                {
                    var cellWbPart = _doc.WorkbookPart
                        ?? throw new InvalidOperationException("Workbook not found");
                    var styleManager = new ExcelStyleManager(cellWbPart);
                    cell.StyleIndex = styleManager.ApplyStyle(cell, cellStyleProps);
                }

                GetSheet(cellWorksheet).Save();
                return $"/{cellSheetName}/{cellRef}";

            case "databar":
            case "conditionalformatting":
            {
                var cfSegments = parentPath.TrimStart('/').Split('/', 2);
                var cfSheetName = cfSegments[0];
                var cfWorksheet = FindWorksheet(cfSheetName)
                    ?? throw new ArgumentException($"Sheet not found: {cfSheetName}");

                var sqref = properties.GetValueOrDefault("sqref", "A1:A10");
                var minVal = properties.GetValueOrDefault("min", "0");
                var maxVal = properties.GetValueOrDefault("max", "1");
                var cfColor = properties.GetValueOrDefault("color", "638EC6");
                var normalizedColor = (cfColor.Length == 6 ? "FF" : "") + cfColor.ToUpperInvariant();

                var cfRule = new ConditionalFormattingRule
                {
                    Type = ConditionalFormatValues.DataBar,
                    Priority = 1
                };
                var dataBar = new DataBar();
                dataBar.Append(new ConditionalFormatValueObject
                {
                    Type = ConditionalFormatValueObjectValues.Number,
                    Val = minVal
                });
                dataBar.Append(new ConditionalFormatValueObject
                {
                    Type = ConditionalFormatValueObjectValues.Number,
                    Val = maxVal
                });
                dataBar.Append(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = normalizedColor });
                cfRule.Append(dataBar);

                var cf = new ConditionalFormatting(cfRule)
                {
                    SequenceOfReferences = new ListValue<StringValue>(
                        sqref.Split(' ').Select(s => new StringValue(s)))
                };

                // Insert after sheetData (or after existing elements)
                var wsElement = GetSheet(cfWorksheet);
                var sheetDataEl = wsElement.GetFirstChild<SheetData>();
                if (sheetDataEl != null)
                    sheetDataEl.InsertAfterSelf(cf);
                else
                    wsElement.Append(cf);

                GetSheet(cfWorksheet).Save();
                return $"/{cfSheetName}/conditionalFormatting[{sqref}]";
            }

            default:
            {
                // Generic fallback: create typed element via SDK schema validation
                // Parse parentPath: /<SheetName>/xmlPath...
                var fbSegments = parentPath.TrimStart('/').Split('/', 2);
                var fbSheetName = fbSegments[0];
                var fbWorksheet = FindWorksheet(fbSheetName);
                if (fbWorksheet == null)
                    throw new ArgumentException($"Sheet not found: {fbSheetName}");

                OpenXmlElement fbParent = GetSheet(fbWorksheet);
                if (fbSegments.Length > 1 && !string.IsNullOrEmpty(fbSegments[1]))
                {
                    var xmlSegments = GenericXmlQuery.ParsePathSegments(fbSegments[1]);
                    fbParent = GenericXmlQuery.NavigateByPath(fbParent!, xmlSegments)
                        ?? throw new ArgumentException($"Parent element not found: {parentPath}");
                }

                var created = GenericXmlQuery.TryCreateTypedElement(fbParent!, type, properties, index);
                if (created == null)
                    throw new ArgumentException($"Schema-invalid element type '{type}' for parent '{parentPath}'. " +
                        "Use raw-set --action append with explicit XML instead.");

                GetSheet(fbWorksheet).Save();

                var siblings = fbParent.ChildElements.Where(e => e.LocalName == created.LocalName).ToList();
                var createdIdx = siblings.IndexOf(created) + 1;
                return $"{parentPath}/{created.LocalName}[{createdIdx}]";
            }
        }
    }

    public void Remove(string path)
    {
        var segments = path.TrimStart('/').Split('/', 2);
        var sheetName = segments[0];

        if (segments.Length == 1)
        {
            // Remove entire sheet
            var workbookPart = _doc.WorkbookPart
                ?? throw new InvalidOperationException("Workbook not found");
            var sheets = GetWorkbook().GetFirstChild<Sheets>();
            var sheet = sheets?.Elements<Sheet>()
                .FirstOrDefault(s => s.Name?.Value?.Equals(sheetName, StringComparison.OrdinalIgnoreCase) == true);
            if (sheet == null)
                throw new ArgumentException($"Sheet not found: {sheetName}");

            var relId = sheet.Id?.Value;
            sheet.Remove();
            if (relId != null)
                workbookPart.DeletePart(workbookPart.GetPartById(relId));
            GetWorkbook().Save();
            return;
        }

        // Remove cell or row
        var cellRef = segments[1];
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");
        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>()
            ?? throw new ArgumentException("Sheet has no data");

        // Check if it's a row reference like row[N]
        var rowMatch = Regex.Match(cellRef, @"^row\[(\d+)\]$");
        if (rowMatch.Success)
        {
            var rowIdx = uint.Parse(rowMatch.Groups[1].Value);
            var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIdx)
                ?? throw new ArgumentException($"Row {rowIdx} not found");
            row.Remove();
        }
        else
        {
            // Cell reference
            var cell = FindCell(sheetData, cellRef)
                ?? throw new ArgumentException($"Cell {cellRef} not found");
            cell.Remove();
        }

        GetSheet(worksheet).Save();
    }

    public string Move(string sourcePath, string? targetParentPath, int? index)
    {
        var segments = sourcePath.TrimStart('/').Split('/', 2);
        var sheetName = segments[0];
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");

        if (segments.Length < 2)
            throw new ArgumentException("Cannot move an entire sheet. Use move on rows or elements within a sheet.");

        var elementRef = segments[1];
        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>()
            ?? throw new ArgumentException("Sheet has no data");

        // Determine target
        string effectiveParentPath;
        SheetData targetSheetData;
        if (string.IsNullOrEmpty(targetParentPath))
        {
            effectiveParentPath = $"/{sheetName}";
            targetSheetData = sheetData;
        }
        else
        {
            effectiveParentPath = targetParentPath;
            var tgtSegments = targetParentPath.TrimStart('/').Split('/', 2);
            var tgtWorksheet = FindWorksheet(tgtSegments[0])
                ?? throw new ArgumentException($"Target sheet not found: {tgtSegments[0]}");
            targetSheetData = GetSheet(tgtWorksheet).GetFirstChild<SheetData>()
                ?? throw new ArgumentException("Target sheet has no data");
        }

        // Find and move the row
        var rowMatch = Regex.Match(elementRef, @"^row\[(\d+)\]$");
        if (rowMatch.Success)
        {
            var rowIdx = uint.Parse(rowMatch.Groups[1].Value);
            var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIdx)
                ?? throw new ArgumentException($"Row {rowIdx} not found");
            row.Remove();

            if (index.HasValue)
            {
                var rows = targetSheetData.Elements<Row>().ToList();
                if (index.Value >= 0 && index.Value < rows.Count)
                    rows[index.Value].InsertBeforeSelf(row);
                else
                    targetSheetData.AppendChild(row);
            }
            else
            {
                targetSheetData.AppendChild(row);
            }

            GetSheet(worksheet).Save();
            var newRows = targetSheetData.Elements<Row>().ToList();
            var newIdx = newRows.IndexOf(row) + 1;
            return $"{effectiveParentPath}/row[{newIdx}]";
        }

        throw new ArgumentException($"Move not supported for: {elementRef}. Supported: row[N]");
    }

    public string CopyFrom(string sourcePath, string targetParentPath, int? index)
    {
        var segments = sourcePath.TrimStart('/').Split('/', 2);
        var sheetName = segments[0];
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");

        if (segments.Length < 2)
            throw new ArgumentException("Cannot copy an entire sheet with --from. Use add --type sheet instead.");

        var elementRef = segments[1];
        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>()
            ?? throw new ArgumentException("Sheet has no data");

        // Find target
        var tgtSegments = targetParentPath.TrimStart('/').Split('/', 2);
        var tgtWorksheet = FindWorksheet(tgtSegments[0])
            ?? throw new ArgumentException($"Target sheet not found: {tgtSegments[0]}");
        var targetSheetData = GetSheet(tgtWorksheet).GetFirstChild<SheetData>()
            ?? throw new ArgumentException("Target sheet has no data");

        // Copy row
        var rowMatch = Regex.Match(elementRef, @"^row\[(\d+)\]$");
        if (rowMatch.Success)
        {
            var rowIdx = uint.Parse(rowMatch.Groups[1].Value);
            var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIdx)
                ?? throw new ArgumentException($"Row {rowIdx} not found");
            var clone = (Row)row.CloneNode(true);

            if (index.HasValue)
            {
                var rows = targetSheetData.Elements<Row>().ToList();
                if (index.Value >= 0 && index.Value < rows.Count)
                    rows[index.Value].InsertBeforeSelf(clone);
                else
                    targetSheetData.AppendChild(clone);
            }
            else
            {
                targetSheetData.AppendChild(clone);
            }

            GetSheet(tgtWorksheet).Save();
            var newRows = targetSheetData.Elements<Row>().ToList();
            var newIdx = newRows.IndexOf(clone) + 1;
            return $"{targetParentPath}/row[{newIdx}]";
        }

        throw new ArgumentException($"Copy not supported for: {elementRef}. Supported: row[N]");
    }

    // ==================== Raw Layer ====================

    public string Raw(string partPath, int? startRow = null, int? endRow = null, HashSet<string>? cols = null)
    {
        var workbookPart = _doc.WorkbookPart;
        if (workbookPart == null) return "(empty)";

        if (partPath == "/" || partPath == "/workbook")
            return workbookPart.Workbook?.OuterXml ?? "(empty)";

        if (partPath == "/styles")
        {
            var styleManager = new ExcelStyleManager(workbookPart);
            return styleManager.EnsureStylesPart().Stylesheet!.OuterXml;
        }

        if (partPath == "/sharedstrings")
        {
            var sst = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            return sst?.SharedStringTable?.OuterXml ?? "(no shared strings)";
        }

        // Drawing part: /SheetName/drawing
        var drawingMatch = Regex.Match(partPath, @"^/(.+)/drawing$");
        if (drawingMatch.Success)
        {
            var drawSheetName = drawingMatch.Groups[1].Value;
            var drawWs = FindWorksheet(drawSheetName)
                ?? throw new ArgumentException($"Sheet not found: {drawSheetName}");
            var dp = drawWs.DrawingsPart
                ?? throw new ArgumentException($"Sheet '{drawSheetName}' has no drawings");
            return dp.WorksheetDrawing!.OuterXml;
        }

        // Chart part: /SheetName/chart[N] or /chart[N]
        var chartMatch = Regex.Match(partPath, @"^/(.+)/chart\[(\d+)\]$");
        if (chartMatch.Success)
        {
            var chartSheetName = chartMatch.Groups[1].Value;
            var chartIdx = int.Parse(chartMatch.Groups[2].Value);
            var chartWs = FindWorksheet(chartSheetName)
                ?? throw new ArgumentException($"Sheet not found: {chartSheetName}");
            var chartPart = GetChartPart(chartWs, chartIdx);
            return chartPart.ChartSpace!.OuterXml;
        }

        // Global chart: /chart[N] — searches all sheets
        var globalChartMatch = Regex.Match(partPath, @"^/chart\[(\d+)\]$");
        if (globalChartMatch.Success)
        {
            var chartIdx = int.Parse(globalChartMatch.Groups[1].Value);
            var chartPart = GetGlobalChartPart(chartIdx);
            return chartPart.ChartSpace!.OuterXml;
        }

        // Try as sheet name
        var sheetName = partPath.TrimStart('/');
        var worksheet = FindWorksheet(sheetName);
        if (worksheet != null)
        {
            if (startRow.HasValue || endRow.HasValue || cols != null)
                return RawSheetWithFilter(worksheet, startRow, endRow, cols);
            return GetSheet(worksheet).OuterXml;
        }

        return $"Unknown part: {partPath}. Available: /workbook, /styles, /sharedstrings, /<SheetName>, /<SheetName>/drawing, /<SheetName>/chart[N], /chart[N]";
    }

    private static string RawSheetWithFilter(WorksheetPart worksheetPart, int? startRow, int? endRow, HashSet<string>? cols)
    {
        var worksheet = GetSheet(worksheetPart);
        var sheetData = worksheet.GetFirstChild<SheetData>();
        if (sheetData == null)
            return worksheet.OuterXml;

        var cloned = (Worksheet)worksheet.CloneNode(true);
        var clonedSheetData = cloned.GetFirstChild<SheetData>()!;
        clonedSheetData.RemoveAllChildren();

        foreach (var row in sheetData.Elements<Row>())
        {
            var rowNum = (int)row.RowIndex!.Value;
            if (startRow.HasValue && rowNum < startRow.Value) continue;
            if (endRow.HasValue && rowNum > endRow.Value) break;

            if (cols != null)
            {
                var filteredRow = (Row)row.CloneNode(false);
                filteredRow.RowIndex = row.RowIndex;
                foreach (var cell in row.Elements<Cell>())
                {
                    var colName = ParseCellReference(cell.CellReference?.Value ?? "A1").Column;
                    if (cols.Contains(colName))
                        filteredRow.AppendChild(cell.CloneNode(true));
                }
                clonedSheetData.AppendChild(filteredRow);
            }
            else
            {
                clonedSheetData.AppendChild(row.CloneNode(true));
            }
        }

        return cloned.OuterXml;
    }

    public void RawSet(string partPath, string xpath, string action, string? xml)
    {
        var workbookPart = _doc.WorkbookPart
            ?? throw new InvalidOperationException("No workbook part");

        OpenXmlPartRootElement rootElement;
        if (partPath is "/" or "/workbook")
        {
            rootElement = workbookPart.Workbook
                ?? throw new InvalidOperationException("No workbook");
        }
        else if (partPath == "/styles")
        {
            var styleManager = new ExcelStyleManager(workbookPart);
            rootElement = styleManager.EnsureStylesPart().Stylesheet!;
        }
        else if (partPath == "/sharedstrings")
        {
            var sst = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault()
                ?? throw new InvalidOperationException("No shared strings");
            rootElement = sst.SharedStringTable!;
        }
        else
        {
            // Drawing part: /SheetName/drawing
            var drawingMatch = Regex.Match(partPath, @"^/(.+)/drawing$");
            if (drawingMatch.Success)
            {
                var drawSheetName = drawingMatch.Groups[1].Value;
                var drawWs = FindWorksheet(drawSheetName)
                    ?? throw new ArgumentException($"Sheet not found: {drawSheetName}");
                var dp = drawWs.DrawingsPart
                    ?? throw new ArgumentException($"Sheet '{drawSheetName}' has no drawings");
                rootElement = dp.WorksheetDrawing!;
            }
            else
            {
            // Chart part: /SheetName/chart[N] or /chart[N]
            var chartMatch = Regex.Match(partPath, @"^/(.+)/chart\[(\d+)\]$");
            if (chartMatch.Success)
            {
                var chartSheetName = chartMatch.Groups[1].Value;
                var chartIdx = int.Parse(chartMatch.Groups[2].Value);
                var chartWs = FindWorksheet(chartSheetName)
                    ?? throw new ArgumentException($"Sheet not found: {chartSheetName}");
                var chartPart = GetChartPart(chartWs, chartIdx);
                rootElement = chartPart.ChartSpace!;
            }
            else
            {
                var globalChartMatch = Regex.Match(partPath, @"^/chart\[(\d+)\]$");
                if (globalChartMatch.Success)
                {
                    var chartIdx = int.Parse(globalChartMatch.Groups[1].Value);
                    var chartPart = GetGlobalChartPart(chartIdx);
                    rootElement = chartPart.ChartSpace!;
                }
                else
                {
                    // Try as sheet name
                    var sheetName = partPath.TrimStart('/');
                    var worksheet = FindWorksheet(sheetName)
                        ?? throw new ArgumentException($"Unknown part: {partPath}. Available: /workbook, /styles, /sharedstrings, /<SheetName>, /<SheetName>/chart[N], /chart[N]");
                    rootElement = GetSheet(worksheet);
                }
            }
            }
        }

        var affected = RawXmlHelper.Execute(rootElement, xpath, action, xml);
        rootElement.Save();
        Console.WriteLine($"raw-set: {affected} element(s) affected");
    }

    public (string RelId, string PartPath) AddPart(string parentPartPath, string partType, Dictionary<string, string>? properties = null)
    {
        var workbookPart = _doc.WorkbookPart
            ?? throw new InvalidOperationException("No workbook part");

        switch (partType.ToLowerInvariant())
        {
            case "chart":
                // Charts go under a worksheet's DrawingsPart
                var sheetName = parentPartPath.TrimStart('/');
                var worksheetPart = FindWorksheet(sheetName)
                    ?? throw new ArgumentException(
                        $"Sheet not found: {sheetName}. Chart must be added under a sheet: add-part <file> /<SheetName> --type chart");

                var drawingsPart = worksheetPart.DrawingsPart
                    ?? worksheetPart.AddNewPart<DocumentFormat.OpenXml.Packaging.DrawingsPart>();

                // Initialize DrawingsPart if new
                if (drawingsPart.WorksheetDrawing == null)
                {
                    drawingsPart.WorksheetDrawing =
                        new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing();
                    drawingsPart.WorksheetDrawing.Save();

                    // Link DrawingsPart to worksheet if not already linked
                    if (GetSheet(worksheetPart).GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Drawing>() == null)
                    {
                        var drawingRelId = worksheetPart.GetIdOfPart(drawingsPart);
                        GetSheet(worksheetPart).Append(
                            new DocumentFormat.OpenXml.Spreadsheet.Drawing { Id = drawingRelId });
                        GetSheet(worksheetPart).Save();
                    }
                }

                var chartPart = drawingsPart.AddNewPart<DocumentFormat.OpenXml.Packaging.ChartPart>();
                var relId = drawingsPart.GetIdOfPart(chartPart);

                // Initialize with minimal valid ChartSpace
                chartPart.ChartSpace = new DocumentFormat.OpenXml.Drawing.Charts.ChartSpace(
                    new DocumentFormat.OpenXml.Drawing.Charts.Chart(
                        new DocumentFormat.OpenXml.Drawing.Charts.PlotArea(
                            new DocumentFormat.OpenXml.Drawing.Charts.Layout()
                        )
                    )
                );
                chartPart.ChartSpace.Save();

                var chartIdx = drawingsPart.ChartParts.ToList().IndexOf(chartPart);
                return (relId, $"/{sheetName}/chart[{chartIdx + 1}]");

            default:
                throw new ArgumentException(
                    $"Unknown part type: {partType}. Supported: chart");
        }
    }

    public List<ValidationError> Validate() => RawXmlHelper.ValidateDocument(_doc);

    public void Dispose() => _doc.Dispose();

    // ==================== Private Helpers ====================

    private static Worksheet GetSheet(WorksheetPart part) =>
        part.Worksheet ?? throw new InvalidOperationException("Corrupt file: worksheet data missing");

    private Workbook GetWorkbook() =>
        _doc.WorkbookPart?.Workbook ?? throw new InvalidOperationException("Corrupt file: workbook missing");

    private List<(string Name, WorksheetPart Part)> GetWorksheets()
    {
        var result = new List<(string, WorksheetPart)>();
        var workbook = _doc.WorkbookPart?.Workbook;
        if (workbook == null) return result;

        var sheets = workbook.GetFirstChild<Sheets>();
        if (sheets == null) return result;

        foreach (var sheet in sheets.Elements<Sheet>())
        {
            var name = sheet.Name?.Value ?? "?";
            var id = sheet.Id?.Value;
            if (id == null) continue;
            var part = (WorksheetPart)_doc.WorkbookPart!.GetPartById(id);
            result.Add((name, part));
        }

        return result;
    }

    private WorksheetPart? FindWorksheet(string sheetName)
    {
        foreach (var (name, part) in GetWorksheets())
        {
            if (name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                return part;
        }
        return null;
    }

    private string GetCellDisplayValue(Cell cell)
    {
        var value = cell.CellValue?.Text ?? "";

        if (cell.DataType?.Value == CellValues.SharedString)
        {
            var sst = _doc.WorkbookPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            if (sst?.SharedStringTable != null && int.TryParse(value, out int idx))
            {
                var item = sst.SharedStringTable.Elements<SharedStringItem>().ElementAtOrDefault(idx);
                return item?.InnerText ?? value;
            }
        }

        // Formula cells without cached value: show the formula
        if (string.IsNullOrEmpty(value) && cell.CellFormula != null
            && !string.IsNullOrEmpty(cell.CellFormula.Text))
        {
            return $"={cell.CellFormula.Text}";
        }

        return value;
    }

    private List<DocumentNode> GetSheetChildNodes(string sheetName, SheetData sheetData, int depth)
    {
        var children = new List<DocumentNode>();
        foreach (var row in sheetData.Elements<Row>())
        {
            var rowIdx = row.RowIndex?.Value ?? 0;
            var rowNode = new DocumentNode
            {
                Path = $"/{sheetName}/row[{rowIdx}]",
                Type = "row",
                ChildCount = row.Elements<Cell>().Count()
            };

            if (depth > 0)
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    rowNode.Children.Add(CellToNode(sheetName, cell));
                }
            }

            children.Add(rowNode);
        }
        return children;
    }

    private DocumentNode CellToNode(string sheetName, Cell cell)
    {
        var cellRef = cell.CellReference?.Value ?? "?";
        var value = GetCellDisplayValue(cell);
        var formula = cell.CellFormula?.Text;
        var type = cell.DataType?.Value.ToString() ?? "Number";

        var node = new DocumentNode
        {
            Path = $"/{sheetName}/{cellRef}",
            Type = "cell",
            Text = value,
            Preview = cellRef
        };

        node.Format["type"] = type;
        if (formula != null) node.Format["formula"] = formula;
        if (string.IsNullOrEmpty(value)) node.Format["empty"] = true;

        return node;
    }

    private DocumentNode GetCellRange(string sheetName, SheetData sheetData, string range, int depth)
    {
        var parts = range.Split(':');
        if (parts.Length != 2)
            throw new ArgumentException($"Invalid range: {range}");

        var (startCol, startRow) = ParseCellReference(parts[0]);
        var (endCol, endRow) = ParseCellReference(parts[1]);

        var node = new DocumentNode
        {
            Path = $"/{sheetName}/{range}",
            Type = "range",
            Preview = range
        };

        foreach (var row in sheetData.Elements<Row>())
        {
            var rowIdx = row.RowIndex?.Value ?? 0;
            if (rowIdx < startRow || rowIdx > endRow) continue;

            foreach (var cell in row.Elements<Cell>())
            {
                var (colName, _) = ParseCellReference(cell.CellReference?.Value ?? "A1");
                var colIdx = ColumnNameToIndex(colName);
                if (colIdx < ColumnNameToIndex(startCol) || colIdx > ColumnNameToIndex(endCol)) continue;

                node.Children.Add(CellToNode(sheetName, cell));
            }
        }

        node.ChildCount = node.Children.Count;
        return node;
    }

    private static Cell? FindCell(SheetData sheetData, string cellRef)
    {
        foreach (var row in sheetData.Elements<Row>())
        {
            foreach (var cell in row.Elements<Cell>())
            {
                if (cell.CellReference?.Value?.Equals(cellRef, StringComparison.OrdinalIgnoreCase) == true)
                    return cell;
            }
        }
        return null;
    }

    private static Cell FindOrCreateCell(SheetData sheetData, string cellRef)
    {
        var (colName, rowIdx) = ParseCellReference(cellRef);

        // Find or create row
        var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIdx);
        if (row == null)
        {
            row = new Row { RowIndex = (uint)rowIdx };
            // Insert in order
            var after = sheetData.Elements<Row>().LastOrDefault(r => (r.RowIndex?.Value ?? 0) < rowIdx);
            if (after != null)
                after.InsertAfterSelf(row);
            else
                sheetData.InsertAt(row, 0);
        }

        // Find or create cell
        var cell = row.Elements<Cell>().FirstOrDefault(c =>
            c.CellReference?.Value?.Equals(cellRef, StringComparison.OrdinalIgnoreCase) == true);
        if (cell == null)
        {
            cell = new Cell { CellReference = cellRef.ToUpperInvariant() };
            // Insert in column order
            var afterCell = row.Elements<Cell>().LastOrDefault(c =>
            {
                var (cn, _) = ParseCellReference(c.CellReference?.Value ?? "A1");
                return ColumnNameToIndex(cn) < ColumnNameToIndex(colName);
            });
            if (afterCell != null)
                afterCell.InsertAfterSelf(cell);
            else
                row.InsertAt(cell, 0);
        }

        return cell;
    }

    // ==================== Selector ====================

    private record CellSelector(string? Sheet, string? Column, string? ValueEquals, string? ValueNotEquals,
        string? ValueContains, bool? HasFormula, bool? IsEmpty, string? TypeEquals);

    private CellSelector ParseCellSelector(string selector)
    {
        string? sheet = null;
        string? column = null;
        string? valueEquals = null;
        string? valueNotEquals = null;
        string? valueContains = null;
        bool? hasFormula = null;
        bool? isEmpty = null;
        string? typeEquals = null;

        // Check for sheet prefix: Sheet1!cell[...]
        var exclIdx = selector.IndexOf('!');
        if (exclIdx > 0)
        {
            sheet = selector[..exclIdx];
            selector = selector[(exclIdx + 1)..];
        }

        // Parse element and attributes: cell[attr=value]
        var match = Regex.Match(selector, @"^(\w+)?(.*)$");
        var element = match.Groups[1].Value;

        // Column filter: e.g., "B" or "cell" in column context
        if (element.Length <= 3 && Regex.IsMatch(element, @"^[A-Z]+$", RegexOptions.IgnoreCase))
        {
            column = element.ToUpperInvariant();
        }

        // Parse attributes
        foreach (Match attrMatch in Regex.Matches(selector, @"\[(\w+)(!?=)([^\]]*)\]"))
        {
            var key = attrMatch.Groups[1].Value.ToLowerInvariant();
            var op = attrMatch.Groups[2].Value;
            var val = attrMatch.Groups[3].Value.Trim('\'', '"');

            switch (key)
            {
                case "value" when op == "=": valueEquals = val; break;
                case "value" when op == "!=": valueNotEquals = val; break;
                case "type": typeEquals = val; break;
                case "formula": hasFormula = val.ToLowerInvariant() != "false"; break;
                case "empty": isEmpty = val.ToLowerInvariant() != "false"; break;
            }
        }

        // :contains() pseudo-selector
        var containsMatch = Regex.Match(selector, @":contains\(['""]?(.+?)['""]?\)");
        if (containsMatch.Success) valueContains = containsMatch.Groups[1].Value;

        // :empty pseudo-selector
        if (selector.Contains(":empty")) isEmpty = true;

        // :has(formula) pseudo-selector
        if (selector.Contains(":has(formula)")) hasFormula = true;

        return new CellSelector(sheet, column, valueEquals, valueNotEquals, valueContains, hasFormula, isEmpty, typeEquals);
    }

    private bool MatchesCellSelector(Cell cell, string sheetName, CellSelector selector)
    {
        // Column filter
        if (selector.Column != null)
        {
            var cellRef = cell.CellReference?.Value ?? "";
            var (colName, _) = ParseCellReference(cellRef);
            if (!colName.Equals(selector.Column, StringComparison.OrdinalIgnoreCase))
                return false;
        }

        var value = GetCellDisplayValue(cell);

        // Value filters
        if (selector.ValueEquals != null && !value.Equals(selector.ValueEquals, StringComparison.OrdinalIgnoreCase))
            return false;
        if (selector.ValueNotEquals != null && value.Equals(selector.ValueNotEquals, StringComparison.OrdinalIgnoreCase))
            return false;
        if (selector.ValueContains != null && !value.Contains(selector.ValueContains, StringComparison.OrdinalIgnoreCase))
            return false;

        // Formula filter
        if (selector.HasFormula == true && cell.CellFormula == null)
            return false;
        if (selector.HasFormula == false && cell.CellFormula != null)
            return false;

        // Empty filter
        if (selector.IsEmpty == true && !string.IsNullOrEmpty(value))
            return false;
        if (selector.IsEmpty == false && string.IsNullOrEmpty(value))
            return false;

        // Type filter
        if (selector.TypeEquals != null)
        {
            var type = cell.DataType?.Value.ToString() ?? "Number";
            if (!type.Equals(selector.TypeEquals, StringComparison.OrdinalIgnoreCase))
                return false;
        }

        return true;
    }

    // ==================== Cell Reference Utils ====================

    private static (string Column, int Row) ParseCellReference(string cellRef)
    {
        var match = Regex.Match(cellRef, @"^([A-Z]+)(\d+)$", RegexOptions.IgnoreCase);
        if (!match.Success) return ("A", 1);
        return (match.Groups[1].Value.ToUpperInvariant(), int.Parse(match.Groups[2].Value));
    }

    private static int ColumnNameToIndex(string col)
    {
        int result = 0;
        foreach (var c in col.ToUpperInvariant())
        {
            result = result * 26 + (c - 'A' + 1);
        }
        return result;
    }

    private static DocumentFormat.OpenXml.Packaging.ChartPart GetChartPart(WorksheetPart worksheetPart, int index)
    {
        var drawingsPart = worksheetPart.DrawingsPart
            ?? throw new ArgumentException("Sheet has no drawings/charts");
        var chartParts = drawingsPart.ChartParts.ToList();
        if (index < 1 || index > chartParts.Count)
            throw new ArgumentException($"Chart index {index} out of range (1..{chartParts.Count})");
        return chartParts[index - 1];
    }

    private DocumentFormat.OpenXml.Packaging.ChartPart GetGlobalChartPart(int index)
    {
        var allCharts = new List<DocumentFormat.OpenXml.Packaging.ChartPart>();
        foreach (var (_, worksheetPart) in GetWorksheets())
        {
            if (worksheetPart.DrawingsPart != null)
                allCharts.AddRange(worksheetPart.DrawingsPart.ChartParts);
        }
        if (allCharts.Count == 0)
            throw new ArgumentException("No charts found in workbook");
        if (index < 1 || index > allCharts.Count)
            throw new ArgumentException($"Chart index {index} out of range (1..{allCharts.Count})");
        return allCharts[index - 1];
    }

    private static string IndexToColumnName(int index)
    {
        var result = "";
        while (index > 0)
        {
            index--;
            result = (char)('A' + index % 26) + result;
            index /= 26;
        }
        return result;
    }
}
