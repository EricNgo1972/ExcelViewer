using Microsoft.Extensions.Options;
using MK.ExcelViewer.Options;

namespace MK.ExcelViewer.Storage;

/// <summary>
/// Expires old workbooks and caps total disk use.
///
/// This is a BackgroundService rather than a sweep-on-access, and the reason is precise: the exact
/// failure we're defending against is "nobody is visiting this app and its disk is filling up" —
/// a scenario in which an on-access sweep is, by construction, dormant.
/// </summary>
public sealed class RetentionSweeper(
    WorkbookStore store,
    IOptions<ExcelViewerOptions> options,
    ILogger<RetentionSweeper> log) : BackgroundService
{
    private readonly ExcelViewerOptions _options = options.Value;

    protected override async Task ExecuteAsync(CancellationToken ct)
    {
        await Task.Delay(TimeSpan.FromSeconds(30), ct);   // don't fight the app's own startup

        using var timer = new PeriodicTimer(TimeSpan.FromMinutes(Math.Max(1, _options.SweepIntervalMinutes)));

        do
        {
            try
            {
                Sweep();
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                // One bad entry must never kill the loop — that would silently disable retention.
                log.LogError(ex, "Retention sweep failed");
            }
        }
        while (await timer.WaitForNextTickAsync(ct));
    }

    private void Sweep()
    {
        var entries = store.Enumerate().ToList();
        if (entries.Count == 0) return;

        var deleted = 0;
        long freed = 0;

        // 1. Age, sliding from last access. A workbook someone still opens stays alive; a
        //    fire-and-forget report from last week evaporates.
        if (_options.RetentionDays > 0)
        {
            var cutoff = DateTime.UtcNow.AddDays(-_options.RetentionDays);

            foreach (var entry in entries.Where(e => e.LastAccessUtc < cutoff).ToList())
            {
                store.Delete(entry.Hash);
                entries.Remove(entry);
                deleted++;
                freed += entry.Bytes;
            }
        }

        // 2. Disk cap, evicting least-recently-accessed first.
        if (_options.MaxTotalBytes > 0)
        {
            var total = entries.Sum(e => e.Bytes);

            foreach (var entry in entries.OrderBy(e => e.LastAccessUtc))
            {
                if (total <= _options.MaxTotalBytes) break;
                store.Delete(entry.Hash);
                total -= entry.Bytes;
                deleted++;
                freed += entry.Bytes;
            }
        }

        if (deleted > 0)
            log.LogInformation("Retention sweep removed {Count} workbook(s), freeing {Mb} MB",
                deleted, freed / (1024 * 1024));
    }
}
