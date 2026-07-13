using System.IO.Compression;
using System.Xml.Linq;
using ClosedXML.Excel;

namespace MK.ExcelViewer.Rendering.Color;

/// <summary>
/// The workbook's twelve theme colour slots, read from <c>xl/theme/theme1.xml</c> in the .xlsx zip.
///
/// Two traps live in this file, and both silently produce wrong colours rather than errors:
///
/// 1. <c>dk1</c>/<c>lt1</c> are normally &lt;a:sysClr&gt; (windowText / window), not &lt;a:srgbClr&gt;.
///    A parser that only reads srgbClr finds nothing for the two most-used slots — black text and
///    white background — and silently falls back. The RGB is on sysClr's <c>lastClr</c> attribute.
///
/// 2. The style-level theme index is NOT the clrScheme document order: slots 0/1 and 2/3 are
///    SWAPPED. theme="0" means Background 1, which is <c>lt1</c> — not dk1. ClosedXML's
///    XLThemeColor enum already encodes the swapped semantics, so <see cref="Resolve"/> maps
///    Background1 → lt1 and Text1 → dk1. Wire it to document order and every white background
///    in every workbook renders black.
/// </summary>
internal sealed class ThemePalette
{
    private readonly Dictionary<string, string> _slots;

    private ThemePalette(Dictionary<string, string> slots) => _slots = slots;

    /// <summary>The modern Office scheme — the last resort when a file carries no readable theme.</summary>
    internal static ThemePalette OfficeDefault { get; } = new(new(StringComparer.Ordinal)
    {
        ["dk1"] = "#000000", ["lt1"] = "#FFFFFF", ["dk2"] = "#44546A", ["lt2"] = "#E7E6E6",
        ["accent1"] = "#4472C4", ["accent2"] = "#ED7D31", ["accent3"] = "#A5A5A5",
        ["accent4"] = "#FFC000", ["accent5"] = "#5B9BD5", ["accent6"] = "#70AD47",
        ["hlink"] = "#0563C1", ["folHlink"] = "#954F72",
    });

    /// <summary>
    /// Reads theme1.xml straight out of the workbook zip — the ground truth for THIS file.
    /// Any slot the file doesn't define falls back to the Office default; an unreadable or absent
    /// theme part yields the Office default wholesale. Never throws: a missing theme is a normal
    /// workbook, not a corrupt one.
    /// </summary>
    internal static ThemePalette FromWorkbookBytes(byte[] bytes)
    {
        try
        {
            using var zip = new ZipArchive(new MemoryStream(bytes, writable: false), ZipArchiveMode.Read);
            var entry = zip.Entries.FirstOrDefault(e =>
                e.FullName.StartsWith("xl/theme/", StringComparison.OrdinalIgnoreCase) &&
                e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));
            if (entry is null) return OfficeDefault;

            using var stream = entry.Open();
            var doc = XDocument.Load(stream);

            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            var scheme = doc.Descendants(a + "clrScheme").FirstOrDefault();
            if (scheme is null) return OfficeDefault;

            var slots = new Dictionary<string, string>(OfficeDefault._slots, StringComparer.Ordinal);

            foreach (var slot in scheme.Elements())
            {
                var name = slot.Name.LocalName;                       // dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink
                var srgb = slot.Element(a + "srgbClr")?.Attribute("val")?.Value;
                var sys = slot.Element(a + "sysClr")?.Attribute("lastClr")?.Value;   // ← trap 1

                var hex = Normalize(srgb ?? sys);
                if (hex is not null) slots[name] = hex;
            }

            return new ThemePalette(slots);
        }
        catch (Exception ex) when (ex is InvalidDataException or System.Xml.XmlException or IOException)
        {
            return OfficeDefault;
        }
    }

    /// <summary>ClosedXML's theme slot → the clrScheme element that actually backs it. Note the 0↔1, 2↔3 swap.</summary>
    internal string? Resolve(XLThemeColor theme)
    {
        var key = theme switch
        {
            XLThemeColor.Background1 => "lt1",   // ← trap 2: Background 1 is the LIGHT slot
            XLThemeColor.Text1 => "dk1",
            XLThemeColor.Background2 => "lt2",
            XLThemeColor.Text2 => "dk2",
            XLThemeColor.Accent1 => "accent1",
            XLThemeColor.Accent2 => "accent2",
            XLThemeColor.Accent3 => "accent3",
            XLThemeColor.Accent4 => "accent4",
            XLThemeColor.Accent5 => "accent5",
            XLThemeColor.Accent6 => "accent6",
            XLThemeColor.Hyperlink => "hlink",
            XLThemeColor.FollowedHyperlink => "folHlink",
            _ => null,
        };

        return key is not null && _slots.TryGetValue(key, out var hex) ? hex : null;
    }

    private static string? Normalize(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return null;
        var s = raw.Trim().TrimStart('#');
        if (s.Length == 8) s = s[2..];                       // ARGB → RGB
        if (s.Length != 6) return null;
        foreach (var c in s) if (!Uri.IsHexDigit(c)) return null;
        return "#" + s.ToUpperInvariant();
    }
}
