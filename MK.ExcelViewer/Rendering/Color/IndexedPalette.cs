namespace MK.ExcelViewer.Rendering.Color;

/// <summary>
/// The legacy 56-entry BIFF colour palette that indexed colours point into. Still turns up in
/// .xlsx files written by older tools and by non-Excel generators.
///
/// Indices 64 (system foreground) and 65 (system background) — and 0x7FFF (automatic) — are NOT
/// black and white. They mean "whatever the system default is", which for us is "inherit". Mapping
/// them to concrete colours is what makes automatic-coloured text render black on a black fill.
/// </summary>
internal static class IndexedPalette
{
    private static readonly string?[] Palette =
    [
        "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF",   //  0- 7
        "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF",   //  8-15
        "#800000", "#008000", "#000080", "#808000", "#800080", "#008080", "#C0C0C0", "#808080",   // 16-23
        "#9999FF", "#993366", "#FFFFCC", "#CCFFFF", "#660066", "#FF8080", "#0066CC", "#CCCCFF",   // 24-31
        "#000080", "#FF00FF", "#FFFF00", "#00FFFF", "#800080", "#800000", "#008080", "#0000FF",   // 32-39
        "#00CCFF", "#CCFFFF", "#CCFFCC", "#FFFF99", "#99CCFF", "#FF99CC", "#CC99FF", "#FFCC99",   // 40-47
        "#3366FF", "#33CCCC", "#99CC00", "#FFCC00", "#FF9900", "#FF6600", "#666699", "#969696",   // 48-55
        "#003366", "#339966", "#003300", "#333300", "#993300", "#993366", "#333399", "#333333",   // 56-63
    ];

    /// <summary>Null means "automatic / inherit" — not a colour.</summary>
    internal static string? TryGet(int index)
    {
        if (index is 64 or 65 or 0x7FFF) return null;    // system foreground / background / automatic
        return index >= 0 && index < Palette.Length ? Palette[index] : null;
    }
}
