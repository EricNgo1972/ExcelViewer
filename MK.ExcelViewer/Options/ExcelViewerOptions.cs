namespace MK.ExcelViewer.Options;

public sealed class ExcelViewerOptions
{
    public const string SectionName = "ExcelViewer";

    /// <summary>
    /// Shared secret for the X-API-Key header on POST /api/workbooks. EMPTY MEANS OPEN — the app
    /// serves anonymously and logs a warning at startup, matching mk-FileConverter's posture.
    /// Set via the environment: ExcelViewer__ApiKey.
    /// </summary>
    public string ApiKey { get; set; } = "";

    /// <summary>
    /// The origin used to build viewUrl/downloadUrl. The POST arrives from another BACKEND, quite
    /// possibly over an internal or loopback address — but the URL we hand back gets opened in a
    /// HUMAN's browser. Those two origins are frequently not the same, and this is the setting that
    /// saves you. Empty falls back to the request's own scheme/host.
    /// </summary>
    public string PublicBaseUrl { get; set; } = "";

    public string StorageRoot { get; set; } = "";

    public long MaxUploadBytes { get; set; } = 32L * 1024 * 1024;

    /// <summary>
    /// Cap on the workbook's UNCOMPRESSED size. This — not MaxUploadBytes, and not the render caps —
    /// is what actually bounds parse memory: ClosedXML loads the entire workbook before we truncate
    /// anything, and peak working set runs roughly 5–6x this figure. It is also the zip-bomb guard,
    /// since a 30 KB archive can declare gigabytes of XML.
    /// 128 MB allows a ~2M-cell workbook; a typical generated report is well under 10 MB.
    /// </summary>
    public long MaxUncompressedBytes { get; set; } = 128L * 1024 * 1024;

    /// <summary>Wall-clock backstop on a pathological workbook so it can't pin a request forever.</summary>
    public int ParseTimeoutSeconds { get; set; } = 60;

    // ── Render caps ──────────────────────────────────────────────────────────────────────────
    // Every cell becomes a real <td>; there is no virtualization, so the BROWSER is the binding
    // constraint, not the server. Measured end-to-end (page load, Chrome, 20-column sheet):
    //
    //     100k cells → 3.2 MB HTML →  8s        200k cells → 6.4 MB → 12s
    //     300k cells → 9.6 MB HTML → 20s        400k cells →  13 MB → 38s
    //
    // It goes superlinear past ~200k — 4x the cells costs nearly 5x the time — so that is where the
    // default sits. Raise MaxCellsPerSheet if your users will genuinely wait; past ~300k they won't.
    // Rows and columns are guard rails against one pathological dimension; cells is what binds.
    public int MaxRowsPerSheet { get; set; } = 50_000;
    public int MaxColumnsPerSheet { get; set; } = 1_024;
    public int MaxCellsPerSheet { get; set; } = 200_000;

    /// <summary>Sliding, from last access. 0 disables age-based expiry.</summary>
    public int RetentionDays { get; set; } = 7;

    /// <summary>
    /// LRU ceiling. This is the guard that actually saves the disk when the calling app develops a
    /// loop and posts ten thousand workbooks in an hour — age-based expiry alone would let the disk
    /// fill for a week first. 0 disables.
    /// </summary>
    public long MaxTotalBytes { get; set; } = 2L * 1024 * 1024 * 1024;

    public int SweepIntervalMinutes { get; set; } = 60;

    public int RenderCacheMinutes { get; set; } = 20;

    /// <summary>
    /// Budget for cached rendered workbooks, in MB — a BYTE budget, not an entry count. A count-based
    /// cache is fine while every workbook is small and lethal the moment one isn't: eight 60 MB
    /// renders is half a gigabyte of process memory held for twenty minutes.
    /// </summary>
    public int RenderCacheBudgetMb { get; set; } = 256;

    /// <summary>Per-IP fixed window on ingest. 0 disables.</summary>
    public int PerIpPerMinute { get; set; } = 60;
}
