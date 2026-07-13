using System.Globalization;

namespace MK.ExcelViewer.Rendering.Color;

/// <summary>
/// ECMA-376 §18.8.19 tint/shade. Excel stores "White, Background 1, Darker 15%" as a theme
/// slot plus a tint of -0.1499 — the tint is applied to the colour's HSL *luminance*, not by
/// scaling RGB. Scaling RGB is the usual shortcut and it desaturates and skews the hue, which
/// is why tinted fills so often come out subtly wrong in Excel-to-HTML converters.
/// </summary>
internal static class TintMath
{
    /// <summary>Applies an ECMA-376 tint (-1..+1) to a "#RRGGBB" colour.</summary>
    internal static string ApplyTint(string hex, double tint)
    {
        if (Math.Abs(tint) < 0.0005) return hex;

        var (r, g, b) = FromHex(hex);
        var (h, l, s) = RgbToHsl(r, g, b);

        // With HLSMAX normalised to 1.0:
        //   shade (tint < 0): scale luminance toward 0
        //   tint  (tint > 0): scale luminance toward 1
        l = tint < 0
            ? l * (1.0 + tint)
            : l * (1.0 - tint) + tint;

        return ToHex(HslToRgb(h, Math.Clamp(l, 0.0, 1.0), s));
    }

    internal static (double R, double G, double B) FromHex(string hex)
    {
        var s = hex.AsSpan().TrimStart('#');
        return (int.Parse(s[..2], NumberStyles.HexNumber, CultureInfo.InvariantCulture) / 255.0,
                int.Parse(s.Slice(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture) / 255.0,
                int.Parse(s.Slice(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture) / 255.0);
    }

    internal static string ToHex((double R, double G, double B) c)
    {
        static int Byte(double v) => (int)Math.Round(Math.Clamp(v, 0, 1) * 255.0, MidpointRounding.AwayFromZero);
        return $"#{Byte(c.R):X2}{Byte(c.G):X2}{Byte(c.B):X2}";
    }

    /// <summary>Hue is returned normalised to [0,1), not degrees.</summary>
    internal static (double H, double L, double S) RgbToHsl(double r, double g, double b)
    {
        var max = Math.Max(r, Math.Max(g, b));
        var min = Math.Min(r, Math.Min(g, b));
        var l = (max + min) / 2.0;

        var d = max - min;
        if (d < 1e-9) return (0.0, l, 0.0);   // achromatic

        var s = l > 0.5 ? d / (2.0 - max - min) : d / (max + min);

        double h;
        if (max == r) h = ((g - b) / d + (g < b ? 6.0 : 0.0)) / 6.0;
        else if (max == g) h = ((b - r) / d + 2.0) / 6.0;
        else h = ((r - g) / d + 4.0) / 6.0;

        return (h, l, s);
    }

    internal static (double R, double G, double B) HslToRgb(double h, double l, double s)
    {
        if (s < 1e-9) return (l, l, l);       // achromatic

        var q = l < 0.5 ? l * (1.0 + s) : l + s - l * s;
        var p = 2.0 * l - q;

        return (Component(p, q, h + 1.0 / 3.0), Component(p, q, h), Component(p, q, h - 1.0 / 3.0));
    }

    private static double Component(double p, double q, double t)
    {
        if (t < 0.0) t += 1.0;
        if (t > 1.0) t -= 1.0;

        if (t < 1.0 / 6.0) return p + (q - p) * 6.0 * t;
        if (t < 1.0 / 2.0) return q;
        if (t < 2.0 / 3.0) return p + (q - p) * (2.0 / 3.0 - t) * 6.0;
        return p;
    }
}
