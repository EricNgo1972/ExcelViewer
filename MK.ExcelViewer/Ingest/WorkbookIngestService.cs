using Microsoft.Extensions.Options;
using MK.ExcelViewer.Options;
using MK.ExcelViewer.Rendering;
using MK.ExcelViewer.Storage;

// ReadOptions() lives beside RenderCache — the caps belong to one place, not two.

namespace MK.ExcelViewer.Ingest;

public enum IngestStatus { Created, AlreadyExisted, TooLarge, Unsupported, Unreadable }

public sealed record IngestResult(
    IngestStatus Status,
    string? Hash = null,
    WorkbookMeta? Meta = null,
    string? Error = null)
{
    public bool Ok => Status is IngestStatus.Created or IngestStatus.AlreadyExisted;
}

/// <summary>
/// Validate → hash → parse → store. Both the API endpoint and the Blazor upload page call THIS —
/// the page does not POST to our own HTTP endpoint over loopback. The endpoint and the page are two
/// thin adapters over one service.
/// </summary>
public sealed class WorkbookIngestService(
    WorkbookStore store,
    WorkbookRenderer renderer,
    RenderCache cache,
    IOptions<ExcelViewerOptions> options,
    ILogger<WorkbookIngestService> log)
{
    private readonly ExcelViewerOptions _options = options.Value;

    public async Task<IngestResult> IngestAsync(
        byte[] bytes, string fileName, string source, string? clientIp, CancellationToken ct)
    {
        if (bytes.Length > _options.MaxUploadBytes)
            return new IngestResult(IngestStatus.TooLarge,
                Error: $"File exceeds the {_options.MaxUploadBytes / (1024 * 1024)} MB limit.");

        // The bytes decide what this is — not the extension, and certainly not the client's
        // Content-Type, which is routinely just application/octet-stream.
        var signature = FileSignature.Detect(bytes);
        if (signature.Verdict != SignatureVerdict.Xlsx)
            return new IngestResult(IngestStatus.Unsupported, Error: signature.Rejection);

        // Check the UNCOMPRESSED size before handing anything to the parser. The upload limit says
        // nothing about this — a small archive can declare a workbook far too large to hold in
        // memory, and ClosedXML would load all of it before our render caps ever got a say.
        var uncompressed = FileSignature.UncompressedSize(bytes);
        if (uncompressed > _options.MaxUncompressedBytes)
        {
            log.LogWarning("Rejected {FileName}: {Mb} MB uncompressed, over the {Limit} MB limit.",
                fileName, uncompressed / (1024 * 1024), _options.MaxUncompressedBytes / (1024 * 1024));

            return new IngestResult(IngestStatus.TooLarge,
                Error: $"This workbook is too large to display — it expands to "
                     + $"{uncompressed / (1024 * 1024)} MB, over the "
                     + $"{_options.MaxUncompressedBytes / (1024 * 1024)} MB limit.");
        }

        var hash = WorkbookStore.ComputeHash(bytes);

        if (store.Exists(hash))
        {
            store.Touch(hash);   // refresh the sliding TTL
            var existing = await store.TryGetMetaAsync(hash, ct);
            if (existing is not null)
                return new IngestResult(IngestStatus.AlreadyExisted, hash, existing);
        }

        // Parse EAGERLY, here, rather than lazily at view time. The caller is a backend that then
        // hands a URL to a human: if the workbook is corrupt it must find out now, when it can log
        // and retry — not by shipping a user to a page that explodes. And it costs nothing, because
        // the render goes straight into the cache, so the user's /view seconds later is a cache hit.
        RenderedWorkbook rendered;
        try
        {
            using var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct);
            timeout.CancelAfter(TimeSpan.FromSeconds(_options.ParseTimeoutSeconds));

            var token = timeout.Token;
            rendered = await Task.Run(() => renderer.Render(bytes, fileName, _options.ReadOptions(), token), token);
        }
        catch (OperationCanceledException) when (!ct.IsCancellationRequested)
        {
            return new IngestResult(IngestStatus.Unreadable,
                Error: $"The workbook took longer than {_options.ParseTimeoutSeconds}s to parse.");
        }
        catch (WorkbookReadException ex)
        {
            return new IngestResult(IngestStatus.Unreadable, Error: ex.Message);
        }
        catch (Exception ex)
        {
            log.LogWarning(ex, "Unexpected failure parsing {FileName}", fileName);
            return new IngestResult(IngestStatus.Unreadable, Error: "The workbook could not be read.");
        }

        var meta = new WorkbookMeta
        {
            Hash = hash,
            FileName = fileName,
            SizeBytes = bytes.Length,
            SheetCount = rendered.Sheets.Count,
            SheetNames = rendered.Sheets.Select(s => s.Name).ToList(),
            CreatedUtc = DateTime.UtcNow,
            Source = source,
            ClientIp = clientIp,
        };

        await store.SaveAsync(meta, bytes, ct);
        cache.Set(hash, rendered);

        log.LogInformation("Ingested {FileName} ({Sheets} sheet(s), {Kb} KB) as {Hash}",
            fileName, meta.SheetCount, bytes.Length / 1024, hash);

        return new IngestResult(IngestStatus.Created, hash, meta);
    }
}
