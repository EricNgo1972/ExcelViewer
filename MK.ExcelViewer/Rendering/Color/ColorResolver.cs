using ClosedXML.Excel;

namespace MK.ExcelViewer.Rendering.Color;

/// <summary>
/// XLColor → "#RRGGBB", or null for "automatic / no colour / inherit".
///
/// The trap this exists to close: <c>XLColor.Color</c> THROWS when ColorType is Theme, and Excel's
/// default palette is theme-based — so that property is unusable on most colours in most real
/// workbooks. Wrapping it in a try/catch (the usual hurried fix) silently turns every theme colour
/// into nothing. Always switch on ColorType first; never touch .Color / .Indexed / .ThemeColor
/// without having checked which one is live.
/// </summary>
internal static class ColorResolver
{
    internal static string? ToHex(XLColor? color, ThemePalette theme)
    {
        if (color is null || !color.HasValue) return null;

        switch (color.ColorType)
        {
            case XLColorType.Color:
                var c = color.Color;                                  // System.Drawing.Color (ARGB)
                if (c.A == 0) return null;                            // fully transparent → no colour
                return $"#{c.R:X2}{c.G:X2}{c.B:X2}";                  // alpha is ignored: Excel doesn't blend fills

            case XLColorType.Indexed:
                return IndexedPalette.TryGet(color.Indexed);

            case XLColorType.Theme:
                var base_ = theme.Resolve(color.ThemeColor);
                // Never invent a colour: an unresolved fill inherits white and looks fine,
                // a guessed one looks broken.
                return base_ is null ? null : TintMath.ApplyTint(base_, color.ThemeTint);

            default:
                return null;
        }
    }
}
