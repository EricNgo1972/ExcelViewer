# MK.ExcelViewer

**Your app generates an Excel file. This app shows it in a browser.**

You POST an `.xlsx`, you get back a URL, you send the user there. No Excel install, no download, no
plugin. That's the entire product.

```
your app ──POST the .xlsx──▶ ExcelViewer ──returns {viewUrl}──▶ your app redirects the user
```

- Rendered server-side with **ClosedXML → HTML**: fonts, fills, borders, merged cells, number
  formats, frozen panes, multiple sheet tabs.
- **`.xlsx` only.** Legacy `.xls` is rejected with a message telling you to re-save.
- Hosted at `ghcr.io/ericngo1972/excelviewer`, behind the Maple cloudflared tunnel.

---

# Sending a workbook

There are two flows. Use the **watchable** one if a person is waiting — which is almost always.

| | [Watchable](#watchable-open-the-viewer-first-recommended) | [One-shot](#one-shot-simplest) |
|---|---|---|
| Browser opens | **immediately** | after the upload *and* the parse finish |
| User sees | "Receiving 43% of 5.9 MB", then "Opening the workbook…" | nothing, then the workbook |
| Errors | shown on the page | your app must handle them |
| Calls | 2 (create session, upload) | 1 |

A 6 MB report takes a couple of seconds to upload and ~10 to parse. In the one-shot flow the user
stares at nothing for all of it, because the browser doesn't open until it's over.

## Watchable: open the viewer first (recommended)

1. **Create a session** — instant, returns immediately.
2. **Open `viewUrl` in the browser right away.** The page appears and starts reporting progress.
3. **`PUT` the bytes.** The page tracks them landing, then hands off to the workbook by itself.

```bash
# 1. create — returns in milliseconds
curl -X POST -H "Content-Type: application/json" \
     -d '{"fileName":"VatReport.xlsx","sizeBytes":6042000}' \
     https://excel.maplekiosk.ca/api/sessions
# -> { "sessionId": "...", "viewUrl": ".../open/...", "uploadUrl": ".../api/sessions/.../content" }

# 2. open viewUrl in the user's browser NOW

# 3. send the bytes; the page follows along
curl -X PUT --data-binary @VatReport.xlsx \
     -H "X-File-Name: VatReport.xlsx" \
     "$uploadUrl"
```

`sizeBytes` is optional but worth sending: with it the page shows a real percentage, without it just
a moving bar. The upload is raw bytes, not multipart, precisely so the server can count them as they
arrive.

```csharp
/// <summary>
/// Publishes a workbook so the user can WATCH it arrive. Returns the URL to open immediately —
/// before the file has been sent — then streams the bytes into the session behind it.
/// </summary>
public async Task<string?> PublishWatchableAsync(byte[] workbook, string fileName, CancellationToken ct = default)
{
    var origin = _settings.Endpoint.TrimEnd('/');

    // 1. Create the session. Instant — nothing is uploaded yet.
    using var create = await http.PostAsJsonAsync($"{origin}/api/sessions",
        new { fileName, sizeBytes = workbook.LongLength }, ct);

    if (!create.IsSuccessStatusCode) return null;

    var session = await create.Content.ReadFromJsonAsync<SessionResponse>(ct);
    if (session is null) return null;

    // 2. Hand the URL back NOW, and upload in the background. The user is already looking at the
    //    progress page while these bytes are still in flight.
    _ = Task.Run(async () =>
    {
        try
        {
            using var content = new ByteArrayContent(workbook);
            using var req = new HttpRequestMessage(HttpMethod.Put, session.UploadUrl) { Content = content };
            req.Headers.Add("X-File-Name", fileName);
            await http.SendAsync(req, CancellationToken.None);
        }
        catch (Exception ex)
        {
            // The viewer page surfaces the failure to the user on its own; just record it.
            log.LogError(ex, "ExcelViewer upload failed for {File}", fileName);
        }
    }, CancellationToken.None);

    return session.ViewUrl;
}

private sealed record SessionResponse(string SessionId, string ViewUrl, string UploadUrl);
```

```csharp
var viewUrl = await viewer.PublishWatchableAsync(xlsxBytes, "VatReport.xlsx", ct);
if (viewUrl is not null) return Redirect(viewUrl);      // opens instantly; the page does the rest
```

**If anything goes wrong, the page says so** — a corrupt workbook, a Word file, a file too large, an
interrupted upload. You don't have to surface those yourself; the user is already looking at the page
that reports them.

## One-shot: simplest

`POST /api/workbooks` — `multipart/form-data`, one part named **`file`**. **No authentication.**
The call blocks through upload *and* parse, then returns the URL. Fine for small workbooks, or when
no human is waiting.

```bash
curl -F "file=@report.xlsx" https://excel.maplekiosk.ca/api/workbooks
```

```json
{
  "hash":        "88994c09a4fe2389f9d187e9e4215b75d7e6795fdac07d8449dde3b738acb700",
  "fileName":    "report.xlsx",
  "sizeBytes":   8887,
  "sheetCount":  2,
  "sheetNames":  ["Sales", "Notes"],
  "viewUrl":     "https://excel.maplekiosk.ca/view/88994c09…",
  "downloadUrl": "https://excel.maplekiosk.ca/api/workbooks/88994c09…/original",
  "cached":      false,
  "expiresUtc":  "2026-07-20T17:54:12Z"
}
```

Take `viewUrl`, send the user to it. You're done.

`201 Created` (with the same URL in the `Location` header) when the workbook is new;
`200 OK` with `"cached": true` when you've sent those exact bytes before — see
[Sending the same file twice](#sending-the-same-file-twice).

## The C# client

Drop this into the app that generates the workbook. It **never throws** — a `null` return means
"publish failed", so you can fall back to writing the file to disk. Same shape as
`DocConverterServerService` in the MapleKiosk tree, minus the API key.

```csharp
using System.Net.Http.Headers;
using System.Text.Json;

/// <summary>
/// Publishes a generated workbook to MK.ExcelViewer and returns the URL to open in the user's
/// browser. Never throws; null means the publish failed and the caller should fall back.
/// </summary>
public sealed class ExcelViewerClient(
    HttpClient http,
    IOptions<ExcelViewerSettings> settings,
    ILogger<ExcelViewerClient> log)
{
    private readonly ExcelViewerSettings _settings = settings.Value;

    public async Task<string?> PublishAsync(byte[] workbook, string fileName, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(_settings.Endpoint))
        {
            log.LogWarning("ExcelViewer endpoint is not configured.");
            return null;
        }

        try
        {
            using var form = new MultipartFormDataContent();
            var part = new ByteArrayContent(workbook);
            part.Headers.ContentType = new MediaTypeHeaderValue(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

            form.Add(part, "file", fileName);        // ← the part MUST be named "file"

            using var cts = CancellationTokenSource.CreateLinkedTokenSource(ct);
            cts.CancelAfter(TimeSpan.FromSeconds(30));

            using var resp = await http.PostAsync(
                _settings.Endpoint.TrimEnd('/') + "/api/workbooks", form, cts.Token);

            var body = await resp.Content.ReadAsStringAsync(cts.Token);

            if (!resp.IsSuccessStatusCode)
            {
                // The body is always {"error":"…"} — a sentence you can log or show.
                log.LogWarning("ExcelViewer {Status}: {Body}", (int)resp.StatusCode, body);
                return null;
            }

            using var doc = JsonDocument.Parse(body);
            return doc.RootElement.GetProperty("viewUrl").GetString();
        }
        catch (OperationCanceledException) when (!ct.IsCancellationRequested)
        {
            log.LogWarning("ExcelViewer timed out.");
            return null;
        }
        catch (Exception ex)
        {
            log.LogError(ex, "ExcelViewer publish failed.");
            return null;
        }
    }
}

public sealed class ExcelViewerSettings
{
    public const string SectionName = "ExcelViewer";
    public string Endpoint { get; set; } = "";      // https://excel.maplekiosk.ca
}
```

Register it:

```csharp
builder.Services.Configure<ExcelViewerSettings>(
    builder.Configuration.GetSection(ExcelViewerSettings.SectionName));
builder.Services.AddHttpClient<ExcelViewerClient>();
```

```json
// appsettings.json
"ExcelViewer": { "Endpoint": "https://excel.maplekiosk.ca" }
```

## Sending the user to the workbook

```csharp
var viewUrl = await viewer.PublishAsync(xlsxBytes, $"InvoiceRun-{DateTime.Now:yyyy-MM}.xlsx", ct);

if (viewUrl is null)
{
    // Publishing failed — fall back to the file itself.
    return File(xlsxBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "InvoiceRun.xlsx");
}

return Redirect(viewUrl);                                             // ASP.NET web app
// Process.Start(new ProcessStartInfo(viewUrl) { UseShellExecute = true });   // desktop / WinForms
// <a href="@viewUrl" target="_blank">View report</a>                         // Blazor / Razor
```

> **You cannot put the viewer in an `<iframe>`.** The view page sends
> `Content-Security-Policy: frame-ancestors 'none'`, so an embed will render blank. Redirect, or open
> a new tab. If you genuinely need to embed it, change `frame-ancestors` in `Program.cs` to name your
> app's origin — do not set it to `*`.

## Sending the same file twice

Workbooks are identified by the **SHA-256 of their bytes**, not by a name or an id you pass. Posting
identical bytes returns the **same `viewUrl`**, stores nothing new, and answers `200` with
`"cached": true` instead of `201`.

So retries are free and safe. You don't need to track whether you've already published a report —
just publish it. If the content genuinely changed, the hash changes and you get a new URL.

## How long the link lives

`expiresUtc` in the response tells you. Retention is **sliding, 7 days from last view** — a link
someone keeps opening stays alive; a fire-and-forget report from last week evaporates. There's also a
2 GiB LRU ceiling, so a runaway caller can't fill the disk.

If you show the link to a user, it's fine to say "this link works until {expiresUtc}". If you need it
to live longer, raise `RetentionDays` on the server.

---

# Handling failures

Every error is `{"error": "a sentence you can show or log"}`.

| Status | Means | What your app should do |
|---|---|---|
| `400` | The workbook is corrupt, or there was no `file` part | A real bug — log it. The file your app generated cannot be opened. |
| `413` | Bigger than 32 MB | Don't publish; hand the user the file directly. |
| `415` | Not an `.xlsx` | You sent a `.xls`, `.xlsb`, `.ods`, a Word file, a password-protected workbook, or a corrupt zip. Each gets its own message. Re-save as `.xlsx`. |
| `429` | Rate-limited (60/min per IP) | Back off and retry. |
| `401` | Only if an API key is configured server-side and yours is wrong | Set `X-API-Key`. Off by default. |
| *timeout / connection refused* | Viewer is down | Fall back to the file. Never block your own flow on this. |

**A corrupt workbook fails at POST, not at view time.** The file is parsed eagerly, so if it can't be
read, *you* find out — while you can still log and retry — instead of a user landing on a page that
explodes. Nothing is stored unless it parses.

The one rule for callers: **treat publishing as best-effort.** If the viewer is unreachable, give the
user the `.xlsx`. The client above is written that way — it returns `null` and never throws.

---

# Endpoints

| | |
|---|---|
| `POST /api/sessions` | Start a watchable publish. Optional `{fileName, sizeBytes}`. → `201 {sessionId, viewUrl, uploadUrl}`. Instant. |
| `PUT /api/sessions/{id}/content` | The raw bytes (not multipart), `X-File-Name` header. Progress is tracked as they arrive. → `200 {hash, viewUrl}`. |
| `GET /api/sessions/{id}` | Poll a publish: `{stage, percent, receivedBytes, totalBytes, hash, error}`. Stages: `waiting` → `receiving` → `opening` → `ready` \| `failed`. |
| `GET /open/{sessionId}` | The progress page a user opens **before** the file is sent. Redirects to the workbook on its own. |
| `POST /api/workbooks` | One-shot publish. `multipart/form-data`, part named `file`. → `201` (or `200` if cached). |
| `GET /view/{hash}` | The viewer page. Add `?sheet=2` to open a specific sheet. |
| `GET /api/workbooks/{hash}/original` | Download the exact bytes that were posted. |

### One client, both services

**`/api/documents` is served here as an alias for `/api/workbooks`**, and MK.WordViewer (the `.docx`
sibling) serves `/api/workbooks` as an alias for `/api/documents`. Both nouns hit the same handler on
both services — same rate limit, same `X-API-Key` gate — so a single client class works against either
one with only the base URL changing. Pick whichever noun you prefer and use it everywhere.

The only response difference: this service adds `sheetCount` and `sheetNames`, which WordViewer omits.
Everything else — `hash`, `fileName`, `sizeBytes`, `viewUrl`, `downloadUrl`, `cached`, `expiresUtc` —
is identical, as are the session endpoints, status codes, and the `{"error": "…"}` body shape.
| `GET /health` | Liveness — what the container's HEALTHCHECK hits. |
| `GET /api/health` | Deeper: proves the storage root is writable; reports workbook count and bytes used. |

## Access

**There is no authentication on publishing.** Anyone who can reach the host can POST a workbook;
`/view/{hash}` and the download are likewise open, protected only by the unguessable 64-hex-char
content hash (256 bits). Abuse is bounded by the rate limit and the 2 GiB storage cap, and the app
logs a warning on every boot so an open deployment is never a silent one.

To close it, set `ExcelViewer__ApiKey` on the server; callers then send an `X-API-Key` header. The
guard is already in the code (`Security/ApiKeyGuard.cs`) — it's simply unarmed when the key is empty.

---

# What the viewer renders

Faithfully: cell values with their number formats, fonts, bold/italic/underline, font and fill
colours (including theme colours with tint), borders, alignment, text wrapping and rotation, merged
cells, column widths and row heights, frozen panes, hidden rows and columns, hyperlinks, cell
comments, and multiple sheets as tabs. Formulas show their **cached result**, never the formula text.
Pivot tables render fine, because Excel writes their result into the sheet as ordinary styled cells.

**Not rendered:** charts (ClosedXML has no chart API), embedded images, and conditional formatting —
Excel stores CF *rules*, not the resulting colours, so those cells arrive unformatted. Each of these
shows a banner on the page rather than silently disappearing.

It is not pixel-identical to Excel — HTML's text engine isn't DirectWrite, so a wrapped cell may break
a line one word differently — but every cell's text, font, colour, border, merge, alignment and size
is correct.

## Large reports

Every cell becomes a real `<td>` — there is no virtualization — so **the browser is the binding
constraint, not the server.** Measured end-to-end on a 20-column sheet (page load to painted):

| cells | HTML | gzipped | browser | server RSS |
|---|---|---|---|---|
| 100k | 3.2 MB | 390 KB | 8 s | ~200 MB |
| 200k | 6.4 MB | 783 KB | **12 s** | ~300 MB |
| 300k | 9.6 MB | 1.2 MB | 20 s | ~400 MB |
| 400k | 13 MB | 1.6 MB | 38 s | ~400 MB |

It goes superlinear past ~200k — four times the cells costs nearly five times the time — so
`MaxCellsPerSheet` defaults to **200,000** (50,000 rows / 1,024 columns as guard rails on any one
dimension). Raise it if your users will genuinely wait; past ~300k they won't. Beyond the cap the
page renders what it can and says *"showing the first N of M rows"*, with the original one click away.

Responses are gzip/brotli compressed, which is worth roughly **7×** here — the HTML is a few hundred
style classes repeated across hundreds of thousands of near-identical tags, so it is about as
compressible as text gets.

**The caps do not bound parse memory.** ClosedXML loads the entire workbook before we truncate
anything, so a sheet we only ever show 200,000 cells of is still parsed in full. What bounds that is
`MaxUncompressedBytes` (default 128 MB), checked against the *uncompressed* size of the archive
before the parser is handed anything — which doubles as the zip-bomb guard, since a 30 KB file can
declare gigabytes of XML. Peak working set runs roughly 5–6× that figure, so **budget ~1 GB of
container memory** if you allow the full 128 MB.

The app runs **Workstation GC** deliberately (`ServerGarbageCollection=false`). Server GC keeps a
heap per core and only collects under pressure, which for this workload meant the process climbed to
~1.7 GB after a few large reports and never gave it back. This is a low-throughput viewer, not a
high-concurrency API — collecting promptly matters more than keeping per-core heaps warm. Measured:
1.7 GB → 750 MB on the same workload.

---

# Running and deploying

```bash
dotnet run --project MK.ExcelViewer     # http://localhost:5310
dotnet test                             # 47 tests
docker compose up --build               # :8080
```

**To release:** run the **Release Container** workflow (Actions → Release Container → *Run workflow*)
with a version like `1.0.0`. It builds the image, **boots it and asserts `/health` answers before
pushing**, then publishes `ghcr.io/ericngo1972/excelviewer:{version}` and `:latest`. Paste the catalog
row from the run summary into mk-provisioning:

```
ImageRepository=ghcr.io/ericngo1972/excelviewer   ImageTag=1.0.0   InternalPort=8080
```

The provisioner pulls it, mounts a volume at `/data`, adds the tunnel ingress rule and the CNAME.
Containers are never port-exposed — UFW allows SSH only, so the tunnel is the sole ingress.

### Behind cloudflared

Cloudflare terminates TLS at the edge and forwards plain HTTP to the container on loopback. That
matters more here than for most apps, because **the whole product is a URL we hand back**: with no
forwarded headers, `request.Scheme`/`Host` would be `http://localhost:8080` and every `viewUrl` would
be an unreachable loopback link.

So `UseForwardedHeaders()` runs **first**, honouring `X-Forwarded-For`, `-Proto` and `-Host`, with
`KnownNetworks`/`KnownProxies` cleared — the tunnel is the only path in, so no untrusted peer can
spoof them. With that in place `ExcelViewer__PublicBaseUrl` is **optional**; set it only if you front
the app with something that doesn't forward those headers.

The defaults sit safely inside Cloudflare's edge limits: it times out at ~100 s (our parse timeout is
60) and caps bodies at 100 MB (our upload cap is 32).

## Server config

One `ExcelViewer` section; secrets come from the environment with double underscores
(`ExcelViewer__ApiKey`).

| Setting | Default | Notes |
|---|---|---|
| `ApiKey` | `""` | **Empty = publishing is open** (the intended posture). |
| `PublicBaseUrl` | `""` | Origin used to build `viewUrl`. Optional behind cloudflared. |
| `StorageRoot` | *derived* | `.state/workbooks` in dev; `/var/lib/mk-excelviewer/workbooks` in prod (symlinked to `/data`). |
| `MaxUploadBytes` | 32 MB | Size of the .xlsx on the wire. Under Cloudflare's 100 MB edge cap. |
| `MaxUncompressedBytes` | 128 MB | Size once expanded. **This is what bounds parse memory**, and the zip-bomb guard. |
| `MaxCellsPerSheet` | 200,000 | The binding cap — see [Large reports](#large-reports). |
| `MaxRowsPerSheet` / `MaxColumnsPerSheet` | 50,000 / 1,024 | Guard rails on any one dimension. |
| `RetentionDays` | 7 | Sliding, from last access. |
| `MaxTotalBytes` | 2 GiB | LRU ceiling — the guard that saves the disk if a caller ever loops. |
| `ParseTimeoutSeconds` | 60 | |
| `RenderCacheMinutes` / `RenderCacheBudgetMb` | 20 / 256 | Rendered workbooks cached by hash, so switching sheet tabs never re-parses. A **byte** budget, not an entry count — one big report must evict several small ones rather than sit alongside them. |
| `PerIpPerMinute` | 60 | Rate limit on publishing. 0 disables. |

---

# Design notes

Three decisions that look odd until you know why.

**`/view/{hash}` is static SSR, not Interactive Server.** A rendered grid can be megabytes of HTML.
Pushed through a Blazor Server SignalR circuit, its render tree is held in server memory for the life
of the connection and re-pushed on every reconnect — Blazor Server's worst failure mode. As static
SSR it's one gzipped HTTP response, and the browser gets native scrolling, Ctrl+F, zoom and print for
free. Only the upload page is interactive.

**No MudBlazor**, despite it being pinned in 31 other house projects. Its global reset styles `table`,
`td`, `border-collapse` and `font-family` — exactly the elements the renderer emits with its own
generated CSS. The whole interactive surface here is one file picker.

**No SQLite**, despite mk-FileConverter having one. That database backs a durable job queue with state
transitions and crash recovery. Our work finishes inside the request; the only questions we ask are
"does this hash exist" and "what's oldest", which a directory and a file mtime answer. There's no
`202`/jobId/polling handshake either — ClosedXML parses in milliseconds, so the POST just returns the
answer.

Before changing the renderer, read the tests. They pin the tint math against Excel's own colour
swatches, the theme-slot index swap, the 64px default column, and the handful of cases — frozen cells
that have their own fill, `[Red]` negative number formats, merge-covered cells — where a
plausible-looking change silently produces wrong output.
