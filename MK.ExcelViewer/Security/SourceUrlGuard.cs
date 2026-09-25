namespace MK.ExcelViewer.Security;

/// <summary>
/// Decides whether GET /fetch may download from a URL. /fetch makes OUR server issue a request to a
/// URL a browser handed us, so this is an SSRF boundary: only https, only an exact whitelisted host,
/// only the default port, no credentials. Anything else — including an empty whitelist — is refused.
/// The fetch itself must not follow redirects, or a whitelisted host could bounce us anywhere.
/// </summary>
public static class SourceUrlGuard
{
    public static bool TryParse(string? src, IReadOnlyCollection<string> allowedHosts, out Uri uri)
    {
        uri = null!;
        if (allowedHosts.Count == 0 || string.IsNullOrWhiteSpace(src)) return false;
        if (!Uri.TryCreate(src, UriKind.Absolute, out var parsed)) return false;

        if (parsed.Scheme != Uri.UriSchemeHttps || !parsed.IsDefaultPort) return false;
        if (!string.IsNullOrEmpty(parsed.UserInfo)) return false;

        // Exact match only: "files.phoebus.asia" must not also admit "evil-files.phoebus.asia" or
        // "files.phoebus.asia.evil.com".
        if (!allowedHosts.Any(h => string.Equals(h.Trim(), parsed.IdnHost, StringComparison.OrdinalIgnoreCase)))
            return false;

        uri = parsed;
        return true;
    }

    /// <summary>
    /// The name to show and store. The caller's hint wins (Cloudreve passes the real file name — its
    /// signed URL may not end in one); otherwise the URL's last path segment. Never a path.
    /// </summary>
    public static string FileName(string? hint, Uri src, string fallback)
    {
        foreach (var candidate in new[] { hint, Uri.UnescapeDataString(src.Segments.LastOrDefault() ?? "") })
        {
            var name = Path.GetFileName(candidate?.Replace('\\', '/') ?? "");
            if (!string.IsNullOrWhiteSpace(name)) return name;
        }
        return fallback;
    }
}
