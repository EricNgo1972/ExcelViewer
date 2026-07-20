using System.IO.Compression;
using System.Reflection;
using System.Threading.RateLimiting;
using Microsoft.AspNetCore.DataProtection;
using Microsoft.AspNetCore.Http.Features;
using Microsoft.AspNetCore.ResponseCompression;
using Microsoft.AspNetCore.HttpOverrides;
using Microsoft.AspNetCore.Server.Kestrel.Core;
using Microsoft.Extensions.Options;
using MK.ExcelViewer.Components;
using MK.ExcelViewer.Ingest;
using MK.ExcelViewer.Options;
using MK.ExcelViewer.Rendering;
using MK.ExcelViewer.Security;
using MK.ExcelViewer.Sessions;
using MK.ExcelViewer.Storage;

var builder = WebApplication.CreateBuilder(args);

builder.Services.Configure<ExcelViewerOptions>(
    builder.Configuration.GetSection(ExcelViewerOptions.SectionName));

var cfg = builder.Configuration.GetSection(ExcelViewerOptions.SectionName);

// ── State root (the house convention, from mk-FileConverter / MemberList) ────────────────────
// All durable state lives under ONE folder so a single volume mount survives container recreation.
// The Linux prod path is symlinked onto /data in the Dockerfile.
var stateRoot = builder.Environment.IsDevelopment()
    ? Path.Combine(builder.Environment.ContentRootPath, ".state")
    : OperatingSystem.IsLinux()
        ? "/var/lib/mk-excelviewer"
        : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "MK.ExcelViewer");

if (string.IsNullOrWhiteSpace(builder.Configuration["ExcelViewer:StorageRoot"]))
    builder.Configuration["ExcelViewer:StorageRoot"] = Path.Combine(stateRoot, "workbooks");

// Data Protection keys back the antiforgery token on the upload form. Left at their default, they
// live inside the container filesystem and are regenerated on every redeploy — so anyone holding a
// page from the previous image gets an antiforgery failure on their next upload. Pin them under the
// same state root as everything else (which is the mounted /data volume in prod), exactly as
// MemberList does, so they outlive container recreation.
var keyRing = Path.Combine(stateRoot, "DataProtection");
try
{
    Directory.CreateDirectory(keyRing);
}
catch (Exception ex) when (ex is UnauthorizedAccessException or IOException)
{
    // Nearly always means the state root isn't writable — in the fleet, /data was not mounted, or
    // the container isn't running as root. Say so, rather than dying on a bare IOException stack.
    throw new InvalidOperationException(
        $"Cannot write to the state root '{stateRoot}'. In the container this path is a symlink to " +
        $"/data — check the volume is mounted. Running locally? Set ASPNETCORE_ENVIRONMENT=Development " +
        $"to use ./.state instead.", ex);
}

builder.Services.AddDataProtection()
    .SetApplicationName("MK.ExcelViewer")
    .PersistKeysToFileSystem(new DirectoryInfo(keyRing));

var maxUpload = cfg.GetValue("MaxUploadBytes", 32L * 1024 * 1024);

// ── The four size ceilings, three of which have surprising defaults ──────────────────────────
// Kestrel's is ~28.6 MB and Blazor's InputFile is 512 KB (!). The +4 MB of slack on the transport
// limits matters: multipart boundaries and headers make the wire body larger than the file, so
// setting these EXACTLY to maxUpload makes a file at the limit fail down in Kestrel with a raw 413
// instead of reaching our handler and getting a clean JSON one.
var transportLimit = maxUpload + 4L * 1024 * 1024;
builder.Services.Configure<FormOptions>(o => o.MultipartBodyLengthLimit = transportLimit);
builder.Services.Configure<KestrelServerOptions>(o => o.Limits.MaxRequestBodySize = transportLimit);

// ── Behind the cloudflared tunnel ────────────────────────────────────────────────────────────
// The tunnel terminates TLS at Cloudflare's edge and forwards plain HTTP to us on loopback. Without
// this, request.Scheme is "http" and request.Host is "localhost:8080" — so every viewUrl we hand
// back would be an unreachable loopback URL, which is fatal for this app specifically: the whole
// point of the POST is to return a link a HUMAN's browser can open.
//
// XForwardedHost matters as much as Proto here (MemberList and ChatGateway both take it for the
// same reason). KnownNetworks/KnownProxies are cleared because the tunnel is the ONLY path in — UFW
// on the VPS allows SSH and nothing else, so there is no untrusted peer that could spoof these.
builder.Services.Configure<ForwardedHeadersOptions>(o =>
{
    o.ForwardedHeaders = ForwardedHeaders.XForwardedFor
                       | ForwardedHeaders.XForwardedProto
                       | ForwardedHeaders.XForwardedHost;
    o.KnownNetworks.Clear();
    o.KnownProxies.Clear();
});

builder.Services.AddRazorComponents().AddInteractiveServerComponents();

// ── Response compression ────────────────────────────────────────────────────────────────────
// A rendered sheet is the most compressible thing imaginable: a few hundred distinct style classes
// repeated across hundreds of thousands of near-identical <td> tags. Measured on a 400k-cell sheet
// this takes ~28 MB of HTML down to well under a megabyte on the wire. Without it a large report is
// a multi-megabyte download even on a LAN.
// Cloudflare would compress at the edge in production, but that does nothing for direct/LAN access,
// and it is the origin→edge hop that is slowest anyway.
builder.Services.AddResponseCompression(o =>
{
    o.EnableForHttps = true;
    o.MimeTypes = ResponseCompressionDefaults.MimeTypes.Concat(["text/html; charset=utf-8"]);
});
builder.Services.Configure<BrotliCompressionProviderOptions>(o => o.Level = CompressionLevel.Fastest);
builder.Services.Configure<GzipCompressionProviderOptions>(o => o.Level = CompressionLevel.Fastest);

// The cache is bounded by BYTES, not entries: entry count tells you nothing about the thing that
// actually runs out. Each entry is charged its real rendered size (see RenderCache).
builder.Services.AddMemoryCache(o =>
    o.SizeLimit = (long)cfg.GetValue("RenderCacheBudgetMb", 256) * 1024 * 1024);
builder.Services.AddSingleton<ClosedXmlReader>();
builder.Services.AddSingleton<WorkbookRenderer>();
builder.Services.AddSingleton<RenderCache>();
builder.Services.AddSingleton<WorkbookStore>();
builder.Services.AddSingleton<WorkbookIngestService>();
builder.Services.AddSingleton<SessionStore>();
builder.Services.AddHostedService<RetentionSweeper>();

var perIpPerMinute = cfg.GetValue("PerIpPerMinute", 60);
builder.Services.AddRateLimiter(rl =>
{
    rl.RejectionStatusCode = StatusCodes.Status429TooManyRequests;
    rl.AddPolicy("ingest", http =>
    {
        var ip = http.Connection.RemoteIpAddress?.ToString() ?? "unknown";
        if (perIpPerMinute <= 0) return RateLimitPartition.GetNoLimiter(ip);

        return RateLimitPartition.GetFixedWindowLimiter(ip, _ => new FixedWindowRateLimiterOptions
        {
            PermitLimit = perIpPerMinute,
            Window = TimeSpan.FromMinutes(1),
            QueueLimit = 0,
        });
    });
    rl.OnRejected = async (ctx, token) =>
    {
        ctx.HttpContext.Response.StatusCode = StatusCodes.Status429TooManyRequests;
        await ctx.HttpContext.Response.WriteAsJsonAsync(
            new { error = "Too many requests — slow down and retry shortly." }, token);
    };
});

var app = builder.Build();

var version = Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "1.0.0";

Directory.CreateDirectory(app.Configuration["ExcelViewer:StorageRoot"]!);

// App-to-app calls carry no key by design, so this is the normal posture — but an open upload
// endpoint published through the tunnel is still worth one visible line per boot, because it is
// reachable by anyone who knows the hostname. Setting ExcelViewer__ApiKey turns the gate back on
// (and silences this). Never block startup over it.
if (string.IsNullOrWhiteSpace(app.Configuration["ExcelViewer:ApiKey"]))
    app.Logger.LogWarning(
        "ExcelViewer:ApiKey is not set — POST /api/workbooks accepts anonymous uploads from anyone who can reach this host. " +
        "Abuse is bounded only by the rate limit and the storage cap. Set ExcelViewer__ApiKey to require an X-API-Key header.");

app.UseForwardedHeaders();   // first: everything downstream needs the corrected scheme/host
app.UseResponseCompression();

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/error", createScopeForErrors: true);
    app.UseHsts();
}

// ── CSP ─────────────────────────────────────────────────────────────────────────────────────
// We inject HTML derived from an untrusted workbook, so the renderer is the XSS boundary. It
// encodes everything — this is defence in depth. Note script-src has NO 'unsafe-inline': even a
// renderer bug that let an onerror= through could not execute. Blazor is fine with this; its script
// is a file, not an inline block. style-src does need unsafe-inline, because the workbook's
// generated styles are inline by construction.
app.Use(async (ctx, next) =>
{
    ctx.Response.Headers["Content-Security-Policy"] =
        "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; " +
        "img-src 'self' data:; frame-ancestors 'none'; base-uri 'self'";
    ctx.Response.Headers["X-Content-Type-Options"] = "nosniff";
    await next();
});

app.UseStaticFiles();
app.UseRateLimiter();

// ── X-API-Key gate ──────────────────────────────────────────────────────────────────────────
// Only the ingest POST is gated. /view/{hash} and the download are opened by a human's browser,
// which cannot send a header — they are protected by the 256-bit unguessable hash instead.
app.Use(async (ctx, next) =>
{
    var opts = ctx.RequestServices.GetRequiredService<IOptions<ExcelViewerOptions>>().Value;
    var path = ctx.Request.Path;

    // BOTH ingest nouns must be gated. /api/documents is the MK.WordViewer-compatible alias, and an
    // alias that skipped this check would be an unauthenticated way in to the same handler.
    var isIngestPost =
        (path.StartsWithSegments("/api/workbooks") || path.StartsWithSegments("/api/documents"))
        && HttpMethods.IsPost(ctx.Request.Method);

    var authorized =
        path.StartsWithSegments("/api/health") ? ApiKeyGuard.IsAuthorizedForHealth(ctx, opts)
        : isIngestPost ? ApiKeyGuard.IsAuthorized(ctx, opts)
        : true;

    if (!authorized)
    {
        ctx.Response.StatusCode = StatusCodes.Status401Unauthorized;
        await ctx.Response.WriteAsJsonAsync(new { error = "Invalid or missing X-API-Key." });
        return;
    }

    await next();
});

app.UseAntiforgery();

// ── POST /api/workbooks ─────────────────────────────────────────────────────────────────────
// Synchronous: ClosedXML parses in milliseconds, in-process. mk-FileConverter's 202/jobId/polling
// handshake exists because LibreOffice is slow; here it would be pure ceremony. So: 201 with the
// viewUrl in the body and in Location.
// Two nouns, one handler. /api/workbooks is this app's own name; /api/documents is MK.WordViewer's,
// served here so a single client class works against either service with only the base URL changing.
// If you add another alias, add it to the X-API-Key gate above too.
app.MapPost("/api/workbooks", PostWorkbook).RequireRateLimiting("ingest").DisableAntiforgery();
app.MapPost("/api/documents", PostWorkbook).RequireRateLimiting("ingest").DisableAntiforgery();

static async Task<IResult> PostWorkbook(
    HttpRequest request,
    HttpContext http,
    WorkbookIngestService ingest,
    WorkbookStore store,
    IOptions<ExcelViewerOptions> opts,
    CancellationToken ct)
{
    if (!request.HasFormContentType)
        return Results.BadRequest(new { error = "Expected multipart/form-data." });

    var form = await request.ReadFormAsync(ct);
    var file = form.Files["file"] ?? form.Files.FirstOrDefault();
    if (file is null || file.Length == 0)
        return Results.BadRequest(new { error = "No 'file' field in the request." });

    if (file.Length > opts.Value.MaxUploadBytes)
        return Results.Json(new { error = $"File exceeds the {opts.Value.MaxUploadBytes / (1024 * 1024)} MB limit." },
            statusCode: StatusCodes.Status413PayloadTooLarge);

    byte[] bytes;
    await using (var ms = new MemoryStream())
    {
        await file.CopyToAsync(ms, ct);
        bytes = ms.ToArray();
    }

    var fileName = string.IsNullOrWhiteSpace(file.FileName) ? "workbook.xlsx" : Path.GetFileName(file.FileName);
    var ip = http.Connection.RemoteIpAddress?.ToString();

    var result = await ingest.IngestAsync(bytes, fileName, "api", ip, ct);

    return result.Status switch
    {
        IngestStatus.TooLarge => Results.Json(new { error = result.Error }, statusCode: StatusCodes.Status413PayloadTooLarge),
        IngestStatus.Unsupported => Results.Json(new { error = result.Error }, statusCode: StatusCodes.Status415UnsupportedMediaType),
        IngestStatus.Unreadable => Results.BadRequest(new { error = result.Error }),
        _ => Respond(result, http, store, opts.Value),
    };
}

/// <summary>
/// The origin to put in URLs we hand back. PublicBaseUrl wins when set: the request may have arrived
/// over an internal or loopback address, but these URLs get opened in a HUMAN's browser.
/// </summary>
static string PublicBase(HttpContext http, ExcelViewerOptions opts) =>
    string.IsNullOrWhiteSpace(opts.PublicBaseUrl)
        ? $"{http.Request.Scheme}://{http.Request.Host}"
        : opts.PublicBaseUrl.TrimEnd('/');

static IResult Respond(IngestResult result, HttpContext http, WorkbookStore store, ExcelViewerOptions opts)
{
    var meta = result.Meta!;
    var baseUrl = PublicBase(http, opts);

    var body = new
    {
        hash = meta.Hash,
        fileName = meta.FileName,
        sizeBytes = meta.SizeBytes,
        sheetCount = meta.SheetCount,
        sheetNames = meta.SheetNames,
        viewUrl = $"{baseUrl}/view/{meta.Hash}",
        downloadUrl = $"{baseUrl}/api/workbooks/{meta.Hash}/original",
        cached = result.Status == IngestStatus.AlreadyExisted,
        expiresUtc = store.ExpiresUtc(meta.Hash),
    };

    // 201 says "I created a resource". On a hash hit we didn't — so 200, and the caller can tell.
    return result.Status == IngestStatus.Created
        ? Results.Created($"{baseUrl}/view/{meta.Hash}", body)
        : Results.Ok(body);
}

// ── Watchable publish: open the viewer first, then send the file ────────────────────────────
// POST /api/workbooks is synchronous — the caller blocks through the upload AND the parse, and only
// then gets a URL to open. For a big report that means the user stares at nothing for ten seconds.
//
// This flow inverts that: create a session (instant), open the browser at once, and stream the file
// into the session while the page watches it arrive.
//
//   1. POST /api/sessions        -> { sessionId, viewUrl, uploadUrl }   (instant)
//   2. open viewUrl in the browser                                       (user sees progress)
//   3. PUT  uploadUrl  with the bytes                                    (page tracks % as it lands)

app.MapPost("/api/sessions", (SessionRequest? body, HttpContext http, IOptions<ExcelViewerOptions> opts) =>
{
    var store = http.RequestServices.GetRequiredService<SessionStore>();
    var session = store.Create(body?.FileName, body?.SizeBytes);

    var baseUrl = PublicBase(http, opts.Value);

    return Results.Json(new
    {
        sessionId = session.Id,
        viewUrl = $"{baseUrl}/open/{session.Id}",
        uploadUrl = $"{baseUrl}/api/sessions/{session.Id}/content",
    }, statusCode: StatusCodes.Status201Created);
}).RequireRateLimiting("ingest").DisableAntiforgery();

// The raw bytes, streamed. Not multipart: we want to count bytes as they arrive so the page can show
// a real percentage, and Content-Length gives us the denominator for free.
app.MapPut("/api/sessions/{id}/content", async (
    string id,
    HttpRequest request,
    HttpContext http,
    SessionStore sessions,
    WorkbookIngestService ingest,
    IOptions<ExcelViewerOptions> opts,
    CancellationToken ct) =>
{
    var session = sessions.Get(id);
    if (session is null)
        return Results.NotFound(new { error = "Unknown or expired session." });

    var declared = request.ContentLength;
    if (declared > opts.Value.MaxUploadBytes)
    {
        var tooBig = $"File exceeds the {opts.Value.MaxUploadBytes / (1024 * 1024)} MB limit.";
        session.Fail(tooBig);
        return Results.Json(new { error = tooBig }, statusCode: StatusCodes.Status413PayloadTooLarge);
    }

    var fileName = request.Headers["X-File-Name"].ToString();
    if (string.IsNullOrWhiteSpace(fileName)) fileName = "workbook.xlsx";
    fileName = Path.GetFileName(fileName);

    session.BeginReceiving(fileName, declared);

    byte[] bytes;
    try
    {
        // Copy by hand rather than CopyToAsync so each chunk can move the progress bar. 64 KB is
        // small enough that the bar moves smoothly on a slow link and large enough to be free.
        using var buffer = new MemoryStream(capacity: (int)Math.Min(declared ?? 0, 8 * 1024 * 1024));
        var chunk = new byte[64 * 1024];
        long total = 0;

        int read;
        while ((read = await request.Body.ReadAsync(chunk, ct)) > 0)
        {
            total += read;

            // Enforce the cap against what has ACTUALLY arrived — Content-Length is the sender's
            // claim, and a sender that lies about it must not get to stream us an unbounded body.
            if (total > opts.Value.MaxUploadBytes)
            {
                var tooBig = $"File exceeds the {opts.Value.MaxUploadBytes / (1024 * 1024)} MB limit.";
                session.Fail(tooBig);
                return Results.Json(new { error = tooBig }, statusCode: StatusCodes.Status413PayloadTooLarge);
            }

            await buffer.WriteAsync(chunk.AsMemory(0, read), ct);
            session.Progress(total);
        }

        bytes = buffer.ToArray();
    }
    catch (Exception ex) when (ex is IOException or OperationCanceledException)
    {
        session.Fail("The upload was interrupted before the whole file arrived.");
        return Results.BadRequest(new { error = "Upload interrupted." });
    }

    if (bytes.Length == 0)
    {
        session.Fail("The sending app uploaded an empty file.");
        return Results.BadRequest(new { error = "Empty body." });
    }

    // Everything is here; now the slow part the user is waiting on.
    session.Opening();

    var ip = http.Connection.RemoteIpAddress?.ToString();
    var result = await ingest.IngestAsync(bytes, fileName, "session", ip, ct);

    if (!result.Ok)
    {
        session.Fail(result.Error ?? "The workbook could not be opened.");

        return result.Status switch
        {
            IngestStatus.TooLarge => Results.Json(new { error = result.Error }, statusCode: StatusCodes.Status413PayloadTooLarge),
            IngestStatus.Unsupported => Results.Json(new { error = result.Error }, statusCode: StatusCodes.Status415UnsupportedMediaType),
            _ => Results.BadRequest(new { error = result.Error }),
        };
    }

    session.Ready(result.Hash!);

    var baseUrl = PublicBase(http, opts.Value);
    return Results.Ok(new
    {
        hash = result.Hash,
        viewUrl = $"{baseUrl}/view/{result.Hash}",
        cached = result.Status == IngestStatus.AlreadyExisted,
    });
}).RequireRateLimiting("ingest").DisableAntiforgery();

// Lets the SENDING app poll too — useful when it wants to know the publish landed.
app.MapGet("/api/sessions/{id}", (string id, SessionStore sessions) =>
{
    var session = sessions.Get(id);
    if (session is null) return Results.NotFound(new { error = "Unknown or expired session." });

    var s = session.Snapshot();
    return Results.Json(new
    {
        stage = s.Stage.ToString().ToLowerInvariant(),
        fileName = s.FileName,
        receivedBytes = s.ReceivedBytes,
        totalBytes = s.TotalBytes,
        percent = s.Percent,
        hash = s.Hash,
        error = s.Error,
    });
});

// ── GET /api/workbooks/{hash}/original ──────────────────────────────────────────────────────
// Both nouns again, for the same reason as the POST above.
app.MapGet("/api/workbooks/{hash}/original", GetOriginal);
app.MapGet("/api/documents/{hash}/original", GetOriginal);

static async Task<IResult> GetOriginal(string hash, WorkbookStore store, CancellationToken ct)
{
    var meta = await store.TryGetMetaAsync(hash, ct);
    if (meta is null)
        return Results.NotFound(new { error = "No workbook for that hash (unknown or expired)." });

    store.Touch(hash);

    return Results.File(
        store.OpenRead(hash),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        meta.FileName,
        enableRangeProcessing: true);
}

// ── Health ──────────────────────────────────────────────────────────────────────────────────
// Cheap liveness for the proxy, and a deeper one that proves the store is actually writable —
// because a read-only volume is the failure that would otherwise only surface on the next upload.
app.MapGet("/health", () => Results.Json(new { status = "ok", version }));

app.MapGet("/api/health", (WorkbookStore store) =>
{
    var writable = true;
    string? error = null;
    try
    {
        var probe = Path.Combine(store.Root, ".probe");
        File.WriteAllText(probe, "");
        File.Delete(probe);
    }
    catch (Exception ex)
    {
        writable = false;
        error = ex.Message;
    }

    var entries = store.Enumerate().ToList();

    return Results.Json(new
    {
        status = writable ? "ok" : "degraded",
        version,
        storageWritable = writable,
        error,
        workbookCount = entries.Count,
        bytesUsed = entries.Sum(e => e.Bytes),
    }, statusCode: writable ? StatusCodes.Status200OK : StatusCodes.Status503ServiceUnavailable);
});

app.MapRazorComponents<App>().AddInteractiveServerRenderMode();

app.Run();

/// <summary>Optional hints from the sender, so the page can name the file and show a real percentage.</summary>
public sealed record SessionRequest(string? FileName, long? SizeBytes);
