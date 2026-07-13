using System.IO.Compression;

namespace MK.ExcelViewer.Ingest;

public enum SignatureVerdict { Xlsx, Rejected }

/// <summary>What the bytes actually are, and — when we can't render them — why not, in words a caller can log.</summary>
public readonly record struct SignatureResult(SignatureVerdict Verdict, string? Rejection)
{
    public static SignatureResult Ok => new(SignatureVerdict.Xlsx, null);
    public static SignatureResult No(string why) => new(SignatureVerdict.Rejected, why);
}

/// <summary>
/// Content sniffing for uploads, modelled on mk-FileConverter's ContentSniffer.
///
/// The file EXTENSION is never trusted, and neither is the client's Content-Type — callers routinely
/// send application/octet-stream. Both container formats are ambiguous, and both ambiguities are
/// things people really upload: a .docx is also a PK zip, and a password-protected .xlsx is an OLE2
/// wrapper whose extension says xlsx while its bytes say otherwise. That last case is precisely why
/// extension-sniffing fails.
///
/// Every rejection here is a case that would otherwise reach ClosedXML and come back as a stack trace.
/// </summary>
public static class FileSignature
{
    private static ReadOnlySpan<byte> Zip => [0x50, 0x4B, 0x03, 0x04];                            // "PK\x03\x04"
    private static ReadOnlySpan<byte> Ole2 => [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];   // Compound File

    public static SignatureResult Detect(byte[] bytes)
    {
        if (bytes.Length < 8) return SignatureResult.No("File is too small to be a workbook.");

        var head = bytes.AsSpan();

        if (head.StartsWith(Ole2)) return DetectOle2(bytes);
        if (head.StartsWith(Zip)) return DetectZip(bytes);

        return SignatureResult.No("Not a workbook — the file is not an .xlsx (expected a ZIP container).");
    }

    /// <summary>
    /// OLE2 is never something we can render, but the three cases below are worth telling apart:
    /// each one has a different fix on the caller's side.
    /// </summary>
    private static SignatureResult DetectOle2(byte[] bytes)
    {
        // The CFB directory names are UTF-16LE. Rather than parse the whole compound-file structure
        // for an error message we're about to return anyway, look for the stream names directly.
        var text = System.Text.Encoding.Unicode.GetString(bytes, 0, Math.Min(bytes.Length, 8192));

        if (text.Contains("EncryptedPackage", StringComparison.Ordinal))
            return SignatureResult.No("This workbook is password-protected and cannot be displayed.");

        if (text.Contains("Workbook", StringComparison.Ordinal) || text.Contains("Book", StringComparison.Ordinal))
            return SignatureResult.No("Legacy .xls workbooks are not supported — please save as .xlsx.");

        return SignatureResult.No("Not a workbook — the file is an old Office document, but not a spreadsheet.");
    }

    /// <summary>
    /// Total UNCOMPRESSED size of the .xlsx parts.
    ///
    /// This is the honest bound on how much memory a parse will cost, and it is not something the
    /// render caps can protect: ClosedXML loads the whole workbook into memory before we truncate
    /// anything, so a sheet we would only ever show the first 400,000 cells of is still parsed in
    /// full. Measured, peak working set runs roughly 5–6x this number.
    ///
    /// It is also the zip-bomb guard. A 30 KB archive can legitimately declare gigabytes of XML,
    /// and the upload size limit says nothing about that.
    /// </summary>
    public static long UncompressedSize(byte[] bytes)
    {
        try
        {
            using var zip = new ZipArchive(new MemoryStream(bytes, writable: false), ZipArchiveMode.Read);
            return zip.Entries.Sum(e => e.Length);   // declared size, read from the central directory — nothing is inflated
        }
        catch (InvalidDataException)
        {
            return 0;   // not a readable zip; Detect() will reject it with a proper message
        }
    }

    private static SignatureResult DetectZip(byte[] bytes)
    {
        try
        {
            using var zip = new ZipArchive(new MemoryStream(bytes, writable: false), ZipArchiveMode.Read);

            var names = zip.Entries.Select(e => e.FullName).ToList();

            if (names.Any(n => n.Equals("xl/workbook.xml", StringComparison.OrdinalIgnoreCase)))
                return SignatureResult.Ok;

            // .xlsb is an OOXML container too — same PK header, binary parts. ClosedXML cannot read it.
            if (names.Any(n => n.Equals("xl/workbook.bin", StringComparison.OrdinalIgnoreCase)))
                return SignatureResult.No("Binary .xlsb workbooks are not supported — please save as .xlsx.");

            if (names.Any(n => n.StartsWith("word/", StringComparison.OrdinalIgnoreCase)))
                return SignatureResult.No("That's a Word document, not a workbook.");

            if (names.Any(n => n.StartsWith("ppt/", StringComparison.OrdinalIgnoreCase)))
                return SignatureResult.No("That's a PowerPoint file, not a workbook.");

            if (names.Any(n => n.Equals("mimetype", StringComparison.OrdinalIgnoreCase)))
                return SignatureResult.No("OpenDocument (.ods) spreadsheets are not supported — please save as .xlsx.");

            return SignatureResult.No("Not a workbook — the ZIP contains no xl/workbook.xml.");
        }
        catch (InvalidDataException)
        {
            return SignatureResult.No("The file is corrupt — its ZIP container could not be opened.");
        }
    }
}
