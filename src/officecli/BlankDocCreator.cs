// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Presentation;

namespace OfficeCli;

public static class BlankDocCreator
{
    public static void Create(string path)
    {
        var ext = Path.GetExtension(path).ToLowerInvariant();
        switch (ext)
        {
            case ".xlsx":
                CreateExcel(path);
                break;
            case ".docx":
                CreateWord(path);
                break;
            case ".pptx":
                CreatePowerPoint(path);
                break;
            default:
                throw new NotSupportedException($"Unsupported file type: {ext}. Supported: .docx, .xlsx, .pptx");
        }
        Console.WriteLine($"Created: {path}");
    }

    private static void CreateExcel(string path)
    {
        using var doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = doc.AddWorkbookPart();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>("rId1");
        worksheetPart.Worksheet = new Worksheet(new SheetData());
        worksheetPart.Worksheet.Save();

        workbookPart.Workbook = new Workbook(
            new Sheets(
                new Sheet { Id = "rId1", SheetId = 1, Name = "Sheet1" }
            )
        );
        workbookPart.Workbook.Save();
    }

    private static void CreateWord(string path)
    {
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        mainPart.Document.Save();
    }

    private static void CreatePowerPoint(string path)
    {
        using var doc = PresentationDocument.Create(path, PresentationDocumentType.Presentation);
        var presentationPart = doc.AddPresentationPart();

        // Create SlideMaster + SlideLayout (required by spec)
        var slideMasterPart = presentationPart.AddNewPart<DocumentFormat.OpenXml.Packaging.SlideMasterPart>("rId1");
        var slideLayoutPart = slideMasterPart.AddNewPart<DocumentFormat.OpenXml.Packaging.SlideLayoutPart>("rId1");

        // Theme must be under presentationPart, then shared to slideMaster
        var themePart = presentationPart.AddNewPart<DocumentFormat.OpenXml.Packaging.ThemePart>("rId2");
        slideMasterPart.AddPart(themePart);
        themePart.Theme = new DocumentFormat.OpenXml.Drawing.Theme(
            new DocumentFormat.OpenXml.Drawing.ThemeElements(
                new DocumentFormat.OpenXml.Drawing.ColorScheme(
                    new DocumentFormat.OpenXml.Drawing.Dark1Color(new DocumentFormat.OpenXml.Drawing.SystemColor { Val = DocumentFormat.OpenXml.Drawing.SystemColorValues.WindowText, LastColor = "000000" }),
                    new DocumentFormat.OpenXml.Drawing.Light1Color(new DocumentFormat.OpenXml.Drawing.SystemColor { Val = DocumentFormat.OpenXml.Drawing.SystemColorValues.Window, LastColor = "FFFFFF" }),
                    new DocumentFormat.OpenXml.Drawing.Dark2Color(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "44546A" }),
                    new DocumentFormat.OpenXml.Drawing.Light2Color(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "E7E6E6" }),
                    new DocumentFormat.OpenXml.Drawing.Accent1Color(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "4472C4" }),
                    new DocumentFormat.OpenXml.Drawing.Accent2Color(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "ED7D31" }),
                    new DocumentFormat.OpenXml.Drawing.Accent3Color(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "A5A5A5" }),
                    new DocumentFormat.OpenXml.Drawing.Accent4Color(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "FFC000" }),
                    new DocumentFormat.OpenXml.Drawing.Accent5Color(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "5B9BD5" }),
                    new DocumentFormat.OpenXml.Drawing.Accent6Color(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "70AD47" }),
                    new DocumentFormat.OpenXml.Drawing.Hyperlink(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "0563C1" }),
                    new DocumentFormat.OpenXml.Drawing.FollowedHyperlinkColor(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "954F72" })
                ) { Name = "Office" },
                new DocumentFormat.OpenXml.Drawing.FontScheme(
                    new DocumentFormat.OpenXml.Drawing.MajorFont(
                        new DocumentFormat.OpenXml.Drawing.LatinFont { Typeface = "Calibri Light" },
                        new DocumentFormat.OpenXml.Drawing.EastAsianFont { Typeface = "" },
                        new DocumentFormat.OpenXml.Drawing.ComplexScriptFont { Typeface = "" }
                    ),
                    new DocumentFormat.OpenXml.Drawing.MinorFont(
                        new DocumentFormat.OpenXml.Drawing.LatinFont { Typeface = "Calibri" },
                        new DocumentFormat.OpenXml.Drawing.EastAsianFont { Typeface = "" },
                        new DocumentFormat.OpenXml.Drawing.ComplexScriptFont { Typeface = "" }
                    )
                ) { Name = "Office" },
                new DocumentFormat.OpenXml.Drawing.FormatScheme(
                    new DocumentFormat.OpenXml.Drawing.FillStyleList(
                        new DocumentFormat.OpenXml.Drawing.SolidFill(new DocumentFormat.OpenXml.Drawing.SchemeColor { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor }),
                        new DocumentFormat.OpenXml.Drawing.SolidFill(new DocumentFormat.OpenXml.Drawing.SchemeColor { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor }),
                        new DocumentFormat.OpenXml.Drawing.SolidFill(new DocumentFormat.OpenXml.Drawing.SchemeColor { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor })
                    ),
                    new DocumentFormat.OpenXml.Drawing.LineStyleList(
                        new DocumentFormat.OpenXml.Drawing.Outline(new DocumentFormat.OpenXml.Drawing.SolidFill(new DocumentFormat.OpenXml.Drawing.SchemeColor { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor })) { Width = 6350, CapType = DocumentFormat.OpenXml.Drawing.LineCapValues.Flat },
                        new DocumentFormat.OpenXml.Drawing.Outline(new DocumentFormat.OpenXml.Drawing.SolidFill(new DocumentFormat.OpenXml.Drawing.SchemeColor { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor })) { Width = 12700, CapType = DocumentFormat.OpenXml.Drawing.LineCapValues.Flat },
                        new DocumentFormat.OpenXml.Drawing.Outline(new DocumentFormat.OpenXml.Drawing.SolidFill(new DocumentFormat.OpenXml.Drawing.SchemeColor { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor })) { Width = 19050, CapType = DocumentFormat.OpenXml.Drawing.LineCapValues.Flat }
                    ),
                    new DocumentFormat.OpenXml.Drawing.EffectStyleList(
                        new DocumentFormat.OpenXml.Drawing.EffectStyle(new DocumentFormat.OpenXml.Drawing.EffectList()),
                        new DocumentFormat.OpenXml.Drawing.EffectStyle(new DocumentFormat.OpenXml.Drawing.EffectList()),
                        new DocumentFormat.OpenXml.Drawing.EffectStyle(new DocumentFormat.OpenXml.Drawing.EffectList())
                    ),
                    new DocumentFormat.OpenXml.Drawing.BackgroundFillStyleList(
                        new DocumentFormat.OpenXml.Drawing.SolidFill(new DocumentFormat.OpenXml.Drawing.SchemeColor { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor }),
                        new DocumentFormat.OpenXml.Drawing.SolidFill(new DocumentFormat.OpenXml.Drawing.SchemeColor { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor }),
                        new DocumentFormat.OpenXml.Drawing.SolidFill(new DocumentFormat.OpenXml.Drawing.SchemeColor { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor })
                    )
                ) { Name = "Office" }
            )
        ) { Name = "Office Theme" };
        themePart.Theme.Save();

        slideLayoutPart.SlideLayout = new DocumentFormat.OpenXml.Presentation.SlideLayout(
            new DocumentFormat.OpenXml.Presentation.CommonSlideData(
                new DocumentFormat.OpenXml.Presentation.ShapeTree(
                    new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties(
                        new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties { Id = 1, Name = "" },
                        new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeDrawingProperties(),
                        new DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties()
                    ),
                    new DocumentFormat.OpenXml.Presentation.GroupShapeProperties()
                )
            )
        ) { Type = DocumentFormat.OpenXml.Presentation.SlideLayoutValues.Blank };
        slideLayoutPart.SlideLayout.Save();

        // slideLayout needs back-reference to slideMaster
        slideLayoutPart.AddPart(slideMasterPart);

        slideMasterPart.SlideMaster = new DocumentFormat.OpenXml.Presentation.SlideMaster(
            new DocumentFormat.OpenXml.Presentation.CommonSlideData(
                new DocumentFormat.OpenXml.Presentation.ShapeTree(
                    new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties(
                        new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties { Id = 1, Name = "" },
                        new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeDrawingProperties(),
                        new DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties()
                    ),
                    new DocumentFormat.OpenXml.Presentation.GroupShapeProperties()
                )
            ),
            new DocumentFormat.OpenXml.Presentation.ColorMap
            {
                Background1 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Light1,
                Text1 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Dark1,
                Background2 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Light2,
                Text2 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Dark2,
                Accent1 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent1,
                Accent2 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent2,
                Accent3 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent3,
                Accent4 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent4,
                Accent5 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent5,
                Accent6 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent6,
                Hyperlink = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.FollowedHyperlink,
            },
            new DocumentFormat.OpenXml.Presentation.SlideLayoutIdList(
                new DocumentFormat.OpenXml.Presentation.SlideLayoutId { Id = 2147483649, RelationshipId = "rId1" }
            )
        );
        slideMasterPart.SlideMaster.Save();

        presentationPart.Presentation = new DocumentFormat.OpenXml.Presentation.Presentation(
            new DocumentFormat.OpenXml.Presentation.SlideMasterIdList(
                new DocumentFormat.OpenXml.Presentation.SlideMasterId { Id = 2147483648, RelationshipId = "rId1" }
            ),
            new SlideIdList(),
            new SlideSize { Cx = 12192000, Cy = 6858000 },
            new NotesSize { Cx = 6858000, Cy = 9144000 }
        );
        presentationPart.Presentation.Save();
    }
}
