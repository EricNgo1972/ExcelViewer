using ClosedXML.Excel;
using MK.ExcelViewer.Ingest;
using MK.ExcelViewer.Rendering;
using Xunit;

namespace MK.ExcelViewer.Tests;

/// <summary>
/// These drive the renderer through real .xlsx bytes built with ClosedXML, so they exercise the
/// actual parse path rather than a hand-built model.
/// </summary>
public class RenderingTests
{
    private static readonly WorkbookRenderer Renderer = new(new ClosedXmlReader());

    private static byte[] Build(Action<IXLWorksheet> configure, string sheetName = "Sheet1")
    {
        using var workbook = new XLWorkbook();
        configure(workbook.AddWorksheet(sheetName));
        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    private static string RenderSheet(Action<IXLWorksheet> configure) =>
        Renderer.Render(Build(configure), "test.xlsx").Sheets[0].Html;

    [Fact]
    public void PlainValues_Render()
    {
        var html = RenderSheet(ws =>
        {
            ws.Cell(1, 1).Value = "Hello";
            ws.Cell(1, 2).Value = 42;
        });

        Assert.Contains("Hello", html);
        Assert.Contains("42", html);
    }

    [Fact]
    public void Numbers_AreRightAligned_AndText_Left()
    {
        // Excel's "General" alignment, which has no CSS equivalent and must be resolved per value type.
        var html = RenderSheet(ws =>
        {
            ws.Cell(1, 1).Value = "text";
            ws.Cell(1, 2).Value = 42;
        });

        var textCell = html[html.IndexOf(">text<", StringComparison.Ordinal)..];
        Assert.Contains("ar", html);   // the number got the right-align class
        Assert.Contains("al", html);   // the text got the left-align class
        Assert.NotEmpty(textCell);
    }

    [Fact]
    public void MergedCells_EmitSpans_AndSwallowedCellsEmitNoTd()
    {
        // The row must contain exactly ONE td: the anchor with colspan=3. An empty <td> for a
        // swallowed cell would shove the rest of the row sideways.
        var html = RenderSheet(ws =>
        {
            ws.Cell(1, 1).Value = "Title";
            ws.Range(1, 1, 1, 3).Merge();
        });

        var firstRow = html[html.IndexOf("<tr", StringComparison.Ordinal)..html.IndexOf("</tr>", StringComparison.Ordinal)];

        Assert.Contains("colspan=\"3\"", firstRow);
        Assert.Equal(1, CountOccurrences(firstRow, "<td"));
    }

    [Fact]
    public void EmptyMergedAnchor_StillRenders()
    {
        // A merged-but-empty title bar never appears in CellsUsed. If we don't materialise the
        // anchor, the whole merged region vanishes and the grid shifts.
        var html = RenderSheet(ws => ws.Range(1, 1, 1, 4).Merge());
        Assert.Contains("colspan=\"4\"", html);
    }

    [Fact]
    public void ThemeColorFill_ResolvesFromTheFilesOwnTheme_NotAnAssumedDefault()
    {
        // XLColor.Color THROWS on a theme colour, and the lazy try/catch fix silently drops it —
        // so this asserts a theme fill survives to a concrete hex at all.
        //
        // It also pins something subtler. ClosedXML writes the LEGACY Office theme, in which
        // Accent1 is #4F81BD; the modern Office theme's Accent1 is #4472C4. Getting #4F81BD here
        // is proof we parsed theme1.xml out of this actual file. If someone swaps the resolver for
        // a hardcoded palette, this goes #4472C4 and the test catches it.
        var rendered = Renderer.Render(Build(ws =>
        {
            ws.Cell(1, 1).Value = "x";
            ws.Cell(1, 1).Style.Fill.SetBackgroundColor(XLColor.FromTheme(XLThemeColor.Accent1));
        }), "t.xlsx");

        Assert.Contains("#4F81BD", rendered.Css, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void RedNegativeNumberFormat_ColorsTheCell()
    {
        // ClosedXML applies the number format but throws away its [Red] colour section, so without
        // explicit handling every accounting sheet's negatives render black.
        var rendered = Renderer.Render(Build(ws =>
        {
            ws.Cell(1, 1).Value = -1234.5;
            ws.Cell(1, 1).Style.NumberFormat.Format = "#,##0.00;[Red]-#,##0.00";
        }), "t.xlsx");

        Assert.Contains("color:#FF0000", rendered.Sheets[0].Html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void PositiveValue_UnderRedNegativeFormat_IsNotColored()
    {
        var rendered = Renderer.Render(Build(ws =>
        {
            ws.Cell(1, 1).Value = 1234.5;
            ws.Cell(1, 1).Style.NumberFormat.Format = "#,##0.00;[Red]-#,##0.00";
        }), "t.xlsx");

        Assert.DoesNotContain("color:#FF0000", rendered.Sheets[0].Html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Formula_ShowsCachedValue_NeverTheFormulaText()
    {
        var html = RenderSheet(ws =>
        {
            ws.Cell(1, 1).Value = 2;
            ws.Cell(1, 2).Value = 3;
            ws.Cell(1, 3).FormulaA1 = "A1+B1";
        });

        Assert.DoesNotContain("A1+B1", html);
        Assert.Contains(">5<", html);          // the cached result, not the expression
    }

    [Fact]
    public void FrozenRow_BecomesSticky()
    {
        var html = RenderSheet(ws =>
        {
            ws.Cell(1, 1).Value = "Header";
            ws.Cell(2, 1).Value = "Body";
            ws.SheetView.FreezeRows(1);
        });

        Assert.Contains("position:sticky", html);
        // A sticky cell with no fill is transparent and the body scrolls through it.
        Assert.Contains("background:#FFFFFF", html);
    }

    [Fact]
    public void FrozenCell_WithItsOwnFill_KeepsIt_AndIsNotPaintedWhite()
    {
        // Regression. Frozen cells get an inline background so scrolled content can't show through
        // them — but an inline style BEATS the style class, so forcing it unconditionally erased the
        // fill on every frozen header. A blue title bar with white text on it vanished completely,
        // and no unit test caught it because the correct colour was still sitting in the CSS.
        var rendered = Renderer.Render(Build(ws =>
        {
            ws.Cell(1, 1).Value = "Title";
            ws.Cell(1, 1).Style.Fill.SetBackgroundColor(XLColor.FromArgb(0x44, 0x72, 0xC4));
            ws.Cell(2, 1).Value = "Body";
            ws.SheetView.FreezeRows(1);
        }), "t.xlsx");

        var html = rendered.Sheets[0].Html;
        var titleCell = html[html.IndexOf("<td", StringComparison.Ordinal)..html.IndexOf("</td>", StringComparison.Ordinal)];

        Assert.Contains("position:sticky", titleCell);
        Assert.DoesNotContain("background:#FFFFFF", titleCell);   // its own fill must survive
        Assert.Contains("#4472C4", rendered.Css, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void FrozenCell_WithNoFill_IsPaintedWhite_SoContentCannotScrollThrough()
    {
        var html = RenderSheet(ws =>
        {
            ws.Cell(1, 1).Value = "Header";
            ws.Cell(2, 1).Value = "Body";
            ws.SheetView.FreezeRows(1);
        });

        Assert.Contains("background:#FFFFFF", html);
    }

    [Fact]
    public void RotatedText_TurnsTheOppositeWayFromExcelsSign()
    {
        // Excel is counter-clockwise-positive, CSS is clockwise-positive. A 45° rotation in Excel
        // must become rotate(-45deg), or every rotated header leans the wrong way.
        var html = RenderSheet(ws =>
        {
            ws.Cell(1, 1).Value = "Tilted";
            ws.Cell(1, 1).Style.Alignment.TextRotation = 45;
        });

        Assert.Contains("rotate(-45deg)", html);
    }

    [Fact]
    public void HiddenRow_IsOmitted()
    {
        var html = RenderSheet(ws =>
        {
            ws.Cell(1, 1).Value = "visible";
            ws.Cell(2, 1).Value = "secret";
            ws.Row(2).Hide();
        });

        Assert.Contains("visible", html);
        Assert.DoesNotContain("secret", html);
    }

    [Fact]
    public void BoldFont_ReachesTheCss()
    {
        var rendered = Renderer.Render(Build(ws =>
        {
            ws.Cell(1, 1).Value = "Bold";
            ws.Cell(1, 1).Style.Font.Bold = true;
        }), "t.xlsx");

        Assert.Contains("font-weight:700", rendered.Css);
    }

    [Fact]
    public void MultipleSheets_AllRender_InOrder()
    {
        using var workbook = new XLWorkbook();
        workbook.AddWorksheet("Summary").Cell(1, 1).Value = "s";
        workbook.AddWorksheet("Detail").Cell(1, 1).Value = "d";
        using var stream = new MemoryStream();
        workbook.SaveAs(stream);

        var rendered = Renderer.Render(stream.ToArray(), "t.xlsx");

        Assert.Equal(["Summary", "Detail"], rendered.Sheets.Select(s => s.Name));
    }

    // ── Security ─────────────────────────────────────────────────────────────────────────────
    // The renderer's output is injected as raw markup, so it IS the XSS boundary.

    [Fact]
    public void ScriptTagInACell_IsEncoded_NotExecutable()
    {
        var html = RenderSheet(ws => ws.Cell(1, 1).Value = "<script>alert('xss')</script>");

        Assert.DoesNotContain("<script>", html);
        Assert.Contains("&lt;script&gt;", html);
    }

    [Fact]
    public void JavascriptHyperlink_IsDropped_ButTextSurvives()
    {
        var html = RenderSheet(ws =>
        {
            ws.Cell(1, 1).Value = "Click me";
            ws.Cell(1, 1).SetHyperlink(new XLHyperlink("javascript:alert(1)"));
        });

        Assert.DoesNotContain("javascript:", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Click me", html);
    }

    [Fact]
    public void HttpHyperlink_IsKept()
    {
        var html = RenderSheet(ws =>
        {
            ws.Cell(1, 1).Value = "SPC";
            ws.Cell(1, 1).SetHyperlink(new XLHyperlink("https://maplekiosk.ca"));
        });

        Assert.Contains("href=\"https://maplekiosk.ca", html);
        Assert.Contains("rel=\"noopener noreferrer\"", html);
    }

    // ── Signature sniffing ───────────────────────────────────────────────────────────────────

    [Fact]
    public void RealXlsx_IsAccepted() =>
        Assert.Equal(SignatureVerdict.Xlsx, FileSignature.Detect(Build(ws => ws.Cell(1, 1).Value = 1)).Verdict);

    [Fact]
    public void RandomBytes_AreRejected()
    {
        var result = FileSignature.Detect([0x25, 0x50, 0x44, 0x46, 0x2D, 0x31, 0x2E, 0x34]);   // "%PDF-1.4"
        Assert.Equal(SignatureVerdict.Rejected, result.Verdict);
        Assert.NotNull(result.Rejection);
    }

    [Fact]
    public void LegacyXls_IsRejected_WithAnActionableMessage()
    {
        // An OLE2 header with a "Workbook" stream name — what a real .xls looks like at the front.
        var ole2 = new byte[512];
        byte[] header = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
        header.CopyTo(ole2, 0);
        System.Text.Encoding.Unicode.GetBytes("Workbook").CopyTo(ole2, 128);

        var result = FileSignature.Detect(ole2);

        Assert.Equal(SignatureVerdict.Rejected, result.Verdict);
        Assert.Contains(".xlsx", result.Rejection!);   // tells the caller what to do about it
    }

    [Fact]
    public void CorruptZip_IsRejected_NotThrown()
    {
        byte[] corrupt = [0x50, 0x4B, 0x03, 0x04, 0xFF, 0xFF, 0xFF, 0xFF, 0x00, 0x01];
        Assert.Equal(SignatureVerdict.Rejected, FileSignature.Detect(corrupt).Verdict);
    }

    [Fact]
    public void UncompressedSize_SeesPastTheZipCompression()
    {
        // The upload limit measures the ZIP. Parse memory tracks the XML inside it, which for a
        // workbook is ~9x larger — so this is the number that actually bounds memory, and the only
        // thing standing between us and a decompression bomb.
        var bytes = Build(ws =>
        {
            for (var r = 1; r <= 200; r++)
                for (var c = 1; c <= 20; c++)
                    ws.Cell(r, c).Value = $"cell-{r}-{c}";
        });

        var uncompressed = FileSignature.UncompressedSize(bytes);

        Assert.True(uncompressed > bytes.Length,
            $"uncompressed ({uncompressed}) should exceed the zip ({bytes.Length})");
    }

    [Fact]
    public void UncompressedSize_OnGarbage_IsZero_NotAThrow() =>
        Assert.Equal(0, FileSignature.UncompressedSize([0x50, 0x4B, 0x03, 0x04, 0xFF, 0xFF]));

    [Fact]
    public void CellCap_TruncatesTheSheet_AndSaysSo()
    {
        // Past the cap the sheet must be cut off AND report it, so the page can say "showing the
        // first N of M rows" rather than silently presenting a partial report as if it were whole.
        var bytes = Build(ws =>
        {
            for (var r = 1; r <= 500; r++)
                for (var c = 1; c <= 10; c++)
                    ws.Cell(r, c).Value = r * c;
        });

        var rendered = Renderer.Render(bytes, "big.xlsx",
            new ReadOptions { MaxCellsPerSheet = 1_000, MaxRowsPerSheet = 50_000 });

        var sheet = rendered.Sheets[0];
        Assert.True(sheet.RowsTruncated);
        Assert.Equal(500, sheet.TotalRowCount);      // the honest total
        Assert.True(sheet.ShownRowCount < 500);      // but we only rendered part of it
    }

    [Fact]
    public void UnderTheCap_NothingIsTruncated()
    {
        var bytes = Build(ws =>
        {
            for (var r = 1; r <= 50; r++) ws.Cell(r, 1).Value = r;
        });

        var sheet = Renderer.Render(bytes, "small.xlsx").Sheets[0];

        Assert.False(sheet.RowsTruncated);
        Assert.Equal(50, sheet.ShownRowCount);
    }

    [Fact]
    public void CorruptWorkbook_ThrowsWorkbookReadException_NotSomethingRandom()
    {
        byte[] garbage = [0x50, 0x4B, 0x03, 0x04, 0x00, 0x00, 0x00, 0x00];
        Assert.Throws<WorkbookReadException>(() => Renderer.Render(garbage, "bad.xlsx"));
    }

    private static int CountOccurrences(string haystack, string needle)
    {
        var count = 0;
        var i = 0;
        while ((i = haystack.IndexOf(needle, i, StringComparison.Ordinal)) >= 0) { count++; i += needle.Length; }
        return count;
    }
}
