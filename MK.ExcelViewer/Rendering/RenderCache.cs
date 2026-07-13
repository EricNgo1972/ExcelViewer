using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Options;
using MK.ExcelViewer.Options;
using MK.ExcelViewer.Storage;

namespace MK.ExcelViewer.Rendering;

/// <summary>
/// Keeps rendered workbooks in memory, keyed by content hash, so switching sheet tabs never
/// re-parses. Seeded by the eager parse at ingest, which is why the user's first /view is a hit.
///
/// Bounded by entry COUNT rather than bytes: these are fat objects (a 100k-cell sheet is tens of MB
/// of HTML), and re-parsing is cheap, so a small cache is the right trade.
/// </summary>
public sealed class RenderCache(
    IMemoryCache cache,
    WorkbookStore store,
    WorkbookRenderer renderer,
    IOptions<ExcelViewerOptions> options)
{
    private readonly ExcelViewerOptions _options = options.Value;

    public void Set(string hash, RenderedWorkbook rendered) =>
        cache.Set(hash, rendered, new MemoryCacheEntryOptions
        {
            Size = 1,
            SlidingExpiration = TimeSpan.FromMinutes(_options.RenderCacheMinutes),
        });

    /// <summary>Null when the hash is unknown or its workbook has expired out of the store.</summary>
    public async Task<RenderedWorkbook?> GetOrLoadAsync(string hash, CancellationToken ct)
    {
        if (!WorkbookStore.IsValidHash(hash)) return null;

        if (cache.TryGetValue(hash, out RenderedWorkbook? hit) && hit is not null)
            return hit;

        var meta = await store.TryGetMetaAsync(hash, ct);
        if (meta is null) return null;

        var bytes = await store.TryReadBytesAsync(hash, ct);
        if (bytes is null) return null;

        // Parsing is CPU-bound and a Blazor circuit's synchronisation context is single-threaded,
        // so keep it off that thread.
        var rendered = await Task.Run(() => renderer.Render(bytes, meta.FileName, ct: ct), ct);

        Set(hash, rendered);
        return rendered;
    }
}
