using MK.ExcelViewer.Options;

namespace MK.ExcelViewer.Security;

/// <summary>
/// X-API-Key enforcement, ported from mk-FileConverter (Security/ApiKeyGuard.cs) so the MK services
/// all speak the same shared-secret convention.
///
/// Note the deliberate difference in what we GUARD: the converter also gates /api/result, but our
/// equivalents — /view/{hash} and the original download — are opened by a human's BROWSER, which
/// cannot send a header. So only the ingest POST is key-gated; the view and download URLs are
/// protected by the unguessable 64-hex-char content hash (256 bits) instead.
/// </summary>
public static class ApiKeyGuard
{
    private const string Header = "X-API-Key";

    /// <summary>An empty configured key means auth is disabled (dev).</summary>
    public static bool IsAuthorized(HttpContext ctx, ExcelViewerOptions opts)
    {
        if (string.IsNullOrEmpty(opts.ApiKey)) return true;
        return Equal(ctx.Request.Headers[Header].ToString(), opts.ApiKey);
    }

    /// <summary>
    /// Health stays open for liveness probes, but a key that IS presented must match — which lets a
    /// caller's "Test connection" button tell "reachable but unauthorized" (401) apart from
    /// "reachable and authorized" (200).
    /// </summary>
    public static bool IsAuthorizedForHealth(HttpContext ctx, ExcelViewerOptions opts)
    {
        if (string.IsNullOrEmpty(opts.ApiKey)) return true;
        var provided = ctx.Request.Headers[Header].ToString();
        if (string.IsNullOrEmpty(provided)) return true;   // anonymous probe
        return Equal(provided, opts.ApiKey);
    }

    /// <summary>Constant-time compare: a length-and-bail comparison leaks the key a byte at a time.</summary>
    private static bool Equal(string a, string b)
    {
        if (a.Length != b.Length) return false;
        var diff = 0;
        for (var i = 0; i < a.Length; i++) diff |= a[i] ^ b[i];
        return diff == 0;
    }
}
