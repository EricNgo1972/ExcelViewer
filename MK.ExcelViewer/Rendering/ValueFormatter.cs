using System.Text.RegularExpressions;
using ClosedXML.Excel;
using MK.ExcelViewer.Rendering.Model;

namespace MK.ExcelViewer.Rendering;

/// <summary>
/// Turns a cell into display-ready text plus the value kind, and recovers the one thing ClosedXML's
/// formatter throws away: the number format's colour sections.
/// </summary>
internal static partial class ValueFormatter
{
    /// <summary>
    /// A leading colour token in a number-format section: "[Red]-#,##0.00".
    /// Excel also allows [Color1]..[Color56]; those index the legacy palette.
    /// </summary>
    [GeneratedRegex(@"^\s*\[(Black|Blue|Cyan|Green|Magenta|Red|White|Yellow|Color\s?(\d{1,2}))\]",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex ColorToken();

    private static readonly Dictionary<string, string> NamedColors = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Black"] = "#000000", ["Blue"] = "#0000FF", ["Cyan"] = "#00FFFF", ["Green"] = "#008000",
        ["Magenta"] = "#FF00FF", ["Red"] = "#FF0000", ["White"] = "#FFFFFF", ["Yellow"] = "#FFFF00",
    };

    /// <summary>
    /// The handful of built-in number formats that carry a [Red] negative section. For built-ins,
    /// ClosedXML reports an empty Format string and only the id, so the format string has to be
    /// recovered from the id before we can look for a colour in it.
    /// </summary>
    private static readonly Dictionary<int, string> BuiltInFormats = new()
    {
        [37] = "#,##0 ;(#,##0)",
        [38] = "#,##0 ;[Red](#,##0)",
        [39] = "#,##0.00;(#,##0.00)",
        [40] = "#,##0.00;[Red](#,##0.00)",
    };

    internal static (string Text, CellValueKind Kind, string? ColorOverride) Format(IXLCell cell)
    {
        // A formula cell's value of record is its CACHED result. Reading cell.Value on a formula
        // can trigger ClosedXML's calc engine, which is slow and throws on functions it doesn't
        // implement — and we must never display the formula text itself.
        var value = cell.HasFormula ? cell.CachedValue : cell.Value;
        var kind = KindOf(value);

        string text;
        try
        {
            text = cell.GetFormattedString();
        }
        catch (Exception ex) when (ex is FormatException or InvalidOperationException or NotImplementedException)
        {
            // A format string ClosedXML's engine can't handle. The value is still worth showing.
            text = value.ToString() ?? "";
        }

        return (text, kind, ColorFor(cell, value, kind));
    }

    private static CellValueKind KindOf(XLCellValue value) => value.Type switch
    {
        XLDataType.Blank => CellValueKind.Blank,
        XLDataType.Boolean => CellValueKind.Boolean,
        XLDataType.Number => CellValueKind.Number,
        XLDataType.DateTime => CellValueKind.Date,
        XLDataType.TimeSpan => CellValueKind.Date,
        XLDataType.Error => CellValueKind.Error,
        _ => CellValueKind.Text,
    };

    /// <summary>
    /// Picks the number-format section that applies to this value and returns its colour, if it
    /// declares one. Without this, "[Red]-#,##0.00" renders in black and every accounting sheet's
    /// negative numbers look wrong — the format is applied, but its colour is silently dropped.
    /// </summary>
    private static string? ColorFor(IXLCell cell, XLCellValue value, CellValueKind kind)
    {
        if (kind is not (CellValueKind.Number or CellValueKind.Date)) return null;

        var format = cell.Style.NumberFormat.Format;
        if (string.IsNullOrEmpty(format))
            BuiltInFormats.TryGetValue(cell.Style.NumberFormat.NumberFormatId, out format);
        if (string.IsNullOrEmpty(format) || !format.Contains('[')) return null;

        var sections = SplitSections(format);
        if (sections.Count < 2) return null;   // a single section applies to everything; no sign-specific colour

        var number = value.Type == XLDataType.Number ? value.GetNumber() : 0;

        // Excel's section order is: positive ; negative ; zero ; text.
        var index = number < 0 ? 1
                  : number == 0 && sections.Count >= 3 ? 2
                  : 0;

        var match = ColorToken().Match(sections[index]);
        if (!match.Success) return null;

        var token = match.Groups[1].Value;
        if (NamedColors.TryGetValue(token, out var hex)) return hex;

        // [Color17] → the legacy indexed palette.
        return int.TryParse(match.Groups[2].Value, out var paletteIndex)
            ? Color.IndexedPalette.TryGet(paletteIndex)
            : null;
    }

    /// <summary>
    /// Splits on section separators only — a ';' inside a literal "…" or inside a […] token is data,
    /// not a separator.
    /// </summary>
    private static List<string> SplitSections(string format)
    {
        var sections = new List<string>();
        var start = 0;
        var inQuote = false;
        var inBracket = false;

        for (var i = 0; i < format.Length; i++)
        {
            var c = format[i];
            switch (c)
            {
                case '"': inQuote = !inQuote; break;
                case '[' when !inQuote: inBracket = true; break;
                case ']' when !inQuote: inBracket = false; break;
                case '\\': i++; break;                      // escaped char — skip whatever follows
                case ';' when !inQuote && !inBracket:
                    sections.Add(format[start..i]);
                    start = i + 1;
                    break;
            }
        }

        sections.Add(format[start..]);
        return sections;
    }
}
