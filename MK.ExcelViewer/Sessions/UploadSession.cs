namespace MK.ExcelViewer.Sessions;

public enum UploadStage
{
    /// <summary>The session exists, but the sending app hasn't started uploading yet.</summary>
    Waiting,
    Receiving,
    Opening,
    Ready,
    Failed,
}

/// <summary>An immutable snapshot — what the viewer page renders. Never hand out the live session.</summary>
public sealed record SessionSnapshot(
    UploadStage Stage,
    string? FileName,
    long ReceivedBytes,
    long? TotalBytes,
    string? Hash,
    string? Error)
{
    /// <summary>Null when the sender didn't declare a size — the page then shows an indeterminate bar.</summary>
    public int? Percent => TotalBytes is > 0
        ? (int)Math.Clamp(ReceivedBytes * 100 / TotalBytes.Value, 0, 100)
        : null;
}

/// <summary>
/// One in-flight publish. The upload request writes to it from a request thread while the viewer's
/// Blazor circuit reads it from another, so every access goes through the lock and readers only ever
/// get a snapshot.
/// </summary>
public sealed class UploadSession(string id, string? fileName, long? totalBytes)
{
    private readonly object _gate = new();   // System.Threading.Lock is .NET 9; this targets net8.0

    private UploadStage _stage = UploadStage.Waiting;
    private string? _fileName = fileName;
    private long _received;
    private long? _total = totalBytes;
    private string? _hash;
    private string? _error;

    public string Id { get; } = id;
    public DateTime CreatedUtc { get; } = DateTime.UtcNow;
    public DateTime TouchedUtc { get; private set; } = DateTime.UtcNow;

    public SessionSnapshot Snapshot()
    {
        lock (_gate) return new SessionSnapshot(_stage, _fileName, _received, _total, _hash, _error);
    }

    public void BeginReceiving(string? name, long? total)
    {
        lock (_gate)
        {
            _stage = UploadStage.Receiving;
            _fileName = name ?? _fileName;
            _total = total ?? _total;
            _received = 0;
            TouchedUtc = DateTime.UtcNow;
        }
    }

    public void Progress(long received)
    {
        lock (_gate)
        {
            _received = received;
            TouchedUtc = DateTime.UtcNow;
        }
    }

    public void Opening()
    {
        lock (_gate)
        {
            _stage = UploadStage.Opening;
            TouchedUtc = DateTime.UtcNow;
        }
    }

    public void Ready(string hash)
    {
        lock (_gate)
        {
            _stage = UploadStage.Ready;
            _hash = hash;
            TouchedUtc = DateTime.UtcNow;
        }
    }

    public void Fail(string error)
    {
        lock (_gate)
        {
            _stage = UploadStage.Failed;
            _error = error;
            TouchedUtc = DateTime.UtcNow;
        }
    }
}
