using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Options;
using MK.ExcelViewer.Options;
using MK.ExcelViewer.Storage;

namespace MK.ExcelViewer.Rendering;

/// <summary>
/// Keeps rendered workbooks in memory, keyed by content hash, so switching sheet tabs never
/// re-parses. Seeded by the eager parse at ingest, which is why the user's first /view is a hit.
///
/// The budget is in BYTES, not entries. A count-based cache is fine while every workbook is small
/// and lethal the moment one isn't — a large report can render to tens of MB, so "keep the last 8"
/// quietly becomes "hold half a gigabyte for twenty minutes". Charging each entry its real size
/// means one huge workbook evicts several small ones instead of sitting alongside them.
/// </summary>
public sealed class RenderCache(
    IMemoryCache cache,
    WorkbookStore store,
    WorkbookRenderer renderer,
    IOptions<ExcelViewerOptions> options,
    ILogger<RenderCache> log)
{
    private readonly ExcelViewerOptions _options = options.Value;

    public void Set(string hash, RenderedWorkbook rendered)
    {
        var size = rendered.ApproximateBytes;
        var budget = (long)_options.RenderCacheBudgetMb * 1024 * 1024;

        // A single render larger than the whole budget can never be admitted — MemoryCache would
        // reject it anyway, but silently. Skip it explicitly and say so: it means every view of that
        // workbook re-parses, which is a thing worth seeing in the logs rather than guessing at.
        if (size > budget)
        {
            log.LogWarning(
                "Rendered workbook {Hash} is {Mb} MB, larger than the {Budget} MB render cache — it will be re-rendered on every view. Consider raising ExcelViewer:RenderCacheBudgetMb.",
                hash, size / (1024 * 1024), _options.RenderCacheBudgetMb);
            return;
        }

        cache.Set(hash, rendered, new MemoryCacheEntryOptions
        {
            Size = size,
            SlidingExpiration = TimeSpan.FromMinutes(_options.RenderCacheMinutes),
        });
    }

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
        var rendered = await Task.Run(() => renderer.Render(bytes, meta.FileName, _options.ReadOptions(), ct), ct);

        Set(hash, rendered);
        return rendered;
    }
}

internal static class ReadOptionsExtensions
{
    /// <summary>The configured caps, in the shape the renderer wants.</summary>
    internal static ReadOptions ReadOptions(this ExcelViewerOptions o) => new()
    {
        MaxRowsPerSheet = o.MaxRowsPerSheet,
        MaxColumnsPerSheet = o.MaxColumnsPerSheet,
        MaxCellsPerSheet = o.MaxCellsPerSheet,
    };
}
