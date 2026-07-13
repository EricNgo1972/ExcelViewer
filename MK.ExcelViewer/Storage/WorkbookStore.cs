using System.Security.Cryptography;
using System.Text.Json;
using Microsoft.Extensions.Options;
using MK.ExcelViewer.Options;

namespace MK.ExcelViewer.Storage;

public sealed record WorkbookMeta
{
    public required string Hash { get; init; }
    public required string FileName { get; init; }
    public required long SizeBytes { get; init; }
    public required int SheetCount { get; init; }
    public required IReadOnlyList<string> SheetNames { get; init; }
    public required DateTime CreatedUtc { get; init; }
    public required string Source { get; init; }    // "api" | "upload"
    public string? ClientIp { get; init; }
}

/// <summary>
/// A content-addressed workbook store on disk.
///
/// There is deliberately NO SQLite here, unlike mk-FileConverter. The converter needs a database
/// because it runs a durable job queue with state transitions and crash recovery. Our work finishes
/// inside the request, and the only two questions we ever ask are "does hash H exist?"
/// (Directory.Exists) and "what is the oldest entry?" (a file mtime). A schema, a migration story,
/// and a connection lifetime to answer those would be pure ceremony.
/// </summary>
public sealed class WorkbookStore(IOptions<ExcelViewerOptions> options, ILogger<WorkbookStore> log)
{
    private static readonly JsonSerializerOptions Json = new() { WriteIndented = true };

    private readonly ExcelViewerOptions _options = options.Value;

    public string Root => _options.StorageRoot;

    public static string ComputeHash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    /// <summary>
    /// The one place a user-controlled string is about to become a filesystem path. A hex-only check
    /// closes path traversal completely — no amount of "../.." survives it.
    /// </summary>
    public static bool IsValidHash(string? hash) =>
        hash is { Length: 64 } && hash.All(c => c is >= '0' and <= '9' or >= 'a' and <= 'f');

    /// <summary>Sharded two levels deep so no single directory accumulates thousands of entries.</summary>
    private string DirectoryFor(string hash) => Path.Combine(_options.StorageRoot, hash[..2], hash);

    private string MetaPath(string hash) => Path.Combine(DirectoryFor(hash), "meta.json");
    private string DataPath(string hash) => Path.Combine(DirectoryFor(hash), "original.xlsx");

    public bool Exists(string hash) => IsValidHash(hash) && File.Exists(MetaPath(hash));

    public async Task SaveAsync(WorkbookMeta meta, byte[] bytes, CancellationToken ct)
    {
        var dir = DirectoryFor(meta.Hash);
        Directory.CreateDirectory(dir);

        // Write to a temp name and rename. Rename is atomic on both NTFS and ext4, so a concurrent
        // reader can never observe a half-written file behind a hash that already "exists".
        var dataTmp = DataPath(meta.Hash) + ".tmp";
        var metaTmp = MetaPath(meta.Hash) + ".tmp";

        await File.WriteAllBytesAsync(dataTmp, bytes, ct);
        await File.WriteAllTextAsync(metaTmp, JsonSerializer.Serialize(meta, Json), ct);

        File.Move(dataTmp, DataPath(meta.Hash), overwrite: true);
        File.Move(metaTmp, MetaPath(meta.Hash), overwrite: true);
    }

    public async Task<WorkbookMeta?> TryGetMetaAsync(string hash, CancellationToken ct)
    {
        if (!IsValidHash(hash)) return null;

        var path = MetaPath(hash);
        if (!File.Exists(path)) return null;

        try
        {
            return JsonSerializer.Deserialize<WorkbookMeta>(await File.ReadAllTextAsync(path, ct), Json);
        }
        catch (Exception ex) when (ex is JsonException or IOException)
        {
            log.LogWarning(ex, "Unreadable meta.json for {Hash}", hash);
            return null;
        }
    }

    public async Task<byte[]?> TryReadBytesAsync(string hash, CancellationToken ct)
    {
        if (!IsValidHash(hash)) return null;
        var path = DataPath(hash);
        return File.Exists(path) ? await File.ReadAllBytesAsync(path, ct) : null;
    }

    public Stream OpenRead(string hash) => File.OpenRead(DataPath(hash));

    /// <summary>
    /// meta.json's mtime is the last-access clock — one syscall, no read-modify-write, no lock. The
    /// file is otherwise immutable, so its mtime carries no other meaning, and touching it on every
    /// view gives us a sliding TTL for free.
    /// </summary>
    public void Touch(string hash)
    {
        try
        {
            if (IsValidHash(hash) && File.Exists(MetaPath(hash)))
                File.SetLastWriteTimeUtc(MetaPath(hash), DateTime.UtcNow);
        }
        catch (IOException)
        {
            // A failed touch costs a workbook its sliding reprieve. Not worth failing a request over.
        }
    }

    public DateTime LastAccessUtc(string hash) => File.GetLastWriteTimeUtc(MetaPath(hash));

    public DateTime ExpiresUtc(string hash) => _options.RetentionDays <= 0
        ? DateTime.MaxValue
        : LastAccessUtc(hash).AddDays(_options.RetentionDays);

    /// <summary>Every stored entry, newest access first. Used by the sweeper and /api/health.</summary>
    public IEnumerable<(string Hash, DateTime LastAccessUtc, long Bytes)> Enumerate()
    {
        if (!Directory.Exists(_options.StorageRoot)) yield break;

        foreach (var shard in Directory.EnumerateDirectories(_options.StorageRoot))
            foreach (var dir in Directory.EnumerateDirectories(shard))
            {
                var hash = Path.GetFileName(dir);
                var meta = Path.Combine(dir, "meta.json");
                if (!File.Exists(meta)) continue;

                long bytes = 0;
                DateTime accessed;
                try
                {
                    accessed = File.GetLastWriteTimeUtc(meta);
                    foreach (var f in Directory.EnumerateFiles(dir)) bytes += new FileInfo(f).Length;
                }
                catch (IOException) { continue; }

                yield return (hash, accessed, bytes);
            }
    }

    public void Delete(string hash)
    {
        try
        {
            var dir = DirectoryFor(hash);
            if (Directory.Exists(dir)) Directory.Delete(dir, recursive: true);
        }
        catch (IOException ex)
        {
            log.LogWarning(ex, "Could not delete workbook {Hash}", hash);
        }
    }
}
