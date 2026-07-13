namespace MK.ExcelViewer.Rendering.Layout;

/// <summary>
/// Calibri and Cambria don't exist on macOS or Linux. Left to itself the browser substitutes
/// something with different metrics, and every wrapped cell in the sheet breaks its lines
/// somewhere else — so this is a fidelity concern, not a cosmetic one. Carlito and Caladea are
/// metric-compatible clones of Calibri and Cambria, which makes the substitution invisible.
/// </summary>
internal static class FontStacks
{
    private static readonly Dictionary<string, string> Known = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Calibri"] = "'Calibri', Carlito, 'Segoe UI', sans-serif",
        ["Cambria"] = "'Cambria', Caladea, Georgia, serif",
        ["Arial"] = "Arial, Helvetica, 'Liberation Sans', sans-serif",
        ["Helvetica"] = "Helvetica, Arial, 'Liberation Sans', sans-serif",
        ["Times New Roman"] = "'Times New Roman', 'Liberation Serif', Times, serif",
        ["Courier New"] = "'Courier New', 'Liberation Mono', monospace",
        ["Consolas"] = "Consolas, 'Liberation Mono', monospace",
        ["Segoe UI"] = "'Segoe UI', system-ui, sans-serif",
        ["Verdana"] = "Verdana, 'DejaVu Sans', sans-serif",
        ["Tahoma"] = "Tahoma, 'DejaVu Sans', sans-serif",
    };

    internal static string For(string? fontName)
    {
        if (string.IsNullOrWhiteSpace(fontName)) return Known["Calibri"];
        if (Known.TryGetValue(fontName, out var stack)) return stack;

        // Unknown font: quote it and fall back to a generic. Strip quotes so a font name can't
        // break out of the CSS declaration.
        var safe = fontName.Replace("'", "").Replace("\"", "").Replace(";", "").Trim();
        return safe.Length == 0 ? Known["Calibri"] : $"'{safe}', sans-serif";
    }
}
