using ClosedXML.Excel;
using MK.ExcelViewer.Rendering.Color;
using MK.ExcelViewer.Rendering.Layout;
using Xunit;

namespace MK.ExcelViewer.Tests;

/// <summary>
/// These pin the three things that, when wrong, produce output that looks plausible and IS wrong:
/// the tint math, the theme slot swap, and the column-width constants.
/// </summary>
public class TintMathTests
{
    // The two goldens that prove the HLS round-trip. Both are colours Excel's UI offers by name,
    // so they can be checked by hand against Excel's own swatches.
    [Theory]
    [InlineData("#FFFFFF", -0.1499, "#D9D9D9")]   // "White, Background 1, Darker 15%"
    [InlineData("#4472C4", 0.3999, "#8FAADC")]    // "Blue, Accent 1, Lighter 40%"
    [InlineData("#4472C4", -0.2499, "#2F5597")]   // "Blue, Accent 1, Darker 25%"
    [InlineData("#4472C4", -0.4999, "#203864")]   // "Blue, Accent 1, Darker 50%"
    [InlineData("#000000", 0.4999, "#7F7F7F")]    // "Black, Text 1, Lighter 50%"
    public void ApplyTint_MatchesExcelSwatches(string input, double tint, string expected) =>
        Assert.Equal(expected, TintMath.ApplyTint(input, tint));

    [Fact]
    public void ApplyTint_ZeroTint_IsIdentity() =>
        Assert.Equal("#4472C4", TintMath.ApplyTint("#4472C4", 0.0));

    [Fact]
    public void RgbHslRoundTrip_PreservesColor()
    {
        var (h, l, s) = TintMath.RgbToHsl(68 / 255.0, 114 / 255.0, 196 / 255.0);
        Assert.Equal("#4472C4", TintMath.ToHex(TintMath.HslToRgb(h, l, s)));
    }
}

public class ThemePaletteTests
{
    // The swap is the whole point: theme index 0 is Background 1, which is the LIGHT slot (lt1).
    // If this test ever goes red, every white background in every workbook is rendering black.
    [Fact]
    public void Background1_ResolvesToLightSlot_NotDark() =>
        Assert.Equal("#FFFFFF", ThemePalette.OfficeDefault.Resolve(XLThemeColor.Background1));

    [Fact]
    public void Text1_ResolvesToDarkSlot() =>
        Assert.Equal("#000000", ThemePalette.OfficeDefault.Resolve(XLThemeColor.Text1));

    [Fact]
    public void Background2_ResolvesToLt2() =>
        Assert.Equal("#E7E6E6", ThemePalette.OfficeDefault.Resolve(XLThemeColor.Background2));

    [Fact]
    public void Text2_ResolvesToDk2() =>
        Assert.Equal("#44546A", ThemePalette.OfficeDefault.Resolve(XLThemeColor.Text2));

    [Fact]
    public void Accent1_ResolvesToOfficeBlue() =>
        Assert.Equal("#4472C4", ThemePalette.OfficeDefault.Resolve(XLThemeColor.Accent1));

    [Fact]
    public void FromWorkbookBytes_ReadsTheThemeWrittenByClosedXml()
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet("S").Cell(1, 1).Value = "x";
        using var ms = new MemoryStream();
        wb.SaveAs(ms);

        var palette = ThemePalette.FromWorkbookBytes(ms.ToArray());

        // Whatever theme the file carries, the light slot must be light and the dark slot dark —
        // this is the assertion that catches a document-order (unswapped) regression.
        Assert.Equal("#FFFFFF", palette.Resolve(XLThemeColor.Background1));
        Assert.Equal("#000000", palette.Resolve(XLThemeColor.Text1));
    }

    [Fact]
    public void FromWorkbookBytes_GarbageInput_FallsBackToOfficeDefault()
    {
        var palette = ThemePalette.FromWorkbookBytes([1, 2, 3, 4]);
        Assert.Equal("#4472C4", palette.Resolve(XLThemeColor.Accent1));
    }
}

public class ColorResolverTests
{
    [Fact]
    public void ThemeColor_DoesNotThrow_AndResolves()
    {
        // XLColor.Color throws on a theme colour. This test exists to prove we never touch it.
        var color = XLColor.FromTheme(XLThemeColor.Accent1);
        Assert.Equal("#4472C4", ColorResolver.ToHex(color, ThemePalette.OfficeDefault));
    }

    [Fact]
    public void RgbColor_Resolves() =>
        Assert.Equal("#FF0000", ColorResolver.ToHex(XLColor.FromArgb(255, 0, 0), ThemePalette.OfficeDefault));

    [Fact]
    public void NoColor_IsNull() =>
        Assert.Null(ColorResolver.ToHex(XLColor.NoColor, ThemePalette.OfficeDefault));

    [Fact]
    public void Null_IsNull() =>
        Assert.Null(ColorResolver.ToHex(null, ThemePalette.OfficeDefault));

    // 64 = system foreground, 65 = system background. "Automatic", not black and white.
    [Theory]
    [InlineData(64)]
    [InlineData(65)]
    public void SystemIndexedColors_AreInherit_NotConcrete(int index) =>
        Assert.Null(IndexedPalette.TryGet(index));

    [Fact]
    public void IndexedColor_Resolves() => Assert.Equal("#FF0000", IndexedPalette.TryGet(2));
}

public class MetricsTests
{
    // Excel's default column is 8.43 chars and is exactly 64px wide. This one number is what
    // keeps MaxDigitWidth and ColumnPadding honest.
    [Fact]
    public void DefaultColumn_Is64Px() =>
        Assert.Equal(64.0, Metrics.ColumnPx(Metrics.DefaultColumnWidthChars));

    [Fact]
    public void DefaultRow_Is20Px() =>
        Assert.Equal(20.0, Metrics.RowPx(Metrics.DefaultRowHeightPoints));

    [Fact]
    public void ZeroWidthColumn_IsJustPadding() => Assert.Equal(5.0, Metrics.ColumnPx(0));
}
