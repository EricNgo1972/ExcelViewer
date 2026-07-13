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

    /// <summary>Wall-clock backstop on a pathological workbook so it can't pin a request forever.</summary>
    public int ParseTimeoutSeconds { get; set; } = 60;

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

    /// <summary>Deliberately small: rendered workbooks are fat objects and re-parsing is cheap.</summary>
    public int RenderCacheEntries { get; set; } = 8;

    /// <summary>Per-IP fixed window on ingest. 0 disables.</summary>
    public int PerIpPerMinute { get; set; } = 60;
}
