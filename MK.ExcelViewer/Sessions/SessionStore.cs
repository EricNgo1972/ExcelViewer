using System.Collections.Concurrent;
using System.Security.Cryptography;

namespace MK.ExcelViewer.Sessions;

/// <summary>
/// In-flight publishes, in memory only.
///
/// Deliberately not persisted: a session is meaningful for the handful of seconds between the
/// sending app opening the browser and the workbook being ready. If the process restarts mid-upload
/// the upload is dead anyway — there is nothing worth recovering, and the finished workbook lands in
/// the durable content-addressed store regardless.
/// </summary>
public sealed class SessionStore
{
    private readonly ConcurrentDictionary<string, UploadSession> _sessions = new();

    /// <summary>
    /// Sessions live briefly. Anything untouched for this long is abandoned — the sender crashed, or
    /// nobody ever opened the link.
    /// </summary>
    private static readonly TimeSpan Lifetime = TimeSpan.FromMinutes(30);

    public UploadSession Create(string? fileName, long? totalBytes)
    {
        Sweep();

        // The session id is in a URL that a human opens, so it must be unguessable: knowing one id
        // must not let you watch — or hijack — someone else's publish.
        var id = Convert.ToHexString(RandomNumberGenerator.GetBytes(16)).ToLowerInvariant();

        var session = new UploadSession(id, fileName, totalBytes);
        _sessions[id] = session;
        return session;
    }

    public UploadSession? Get(string id) => _sessions.GetValueOrDefault(id);

    private void Sweep()
    {
        var cutoff = DateTime.UtcNow - Lifetime;

        foreach (var (id, session) in _sessions)
            if (session.TouchedUtc < cutoff)
                _sessions.TryRemove(id, out _);
    }
}
