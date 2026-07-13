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

    var authorized =
        path.StartsWithSegments("/api/health") ? ApiKeyGuard.IsAuthorizedForHealth(ctx, opts)
        : path.StartsWithSegments("/api/workbooks") && HttpMethods.IsPost(ctx.Request.Method)
            ? ApiKeyGuard.IsAuthorized(ctx, opts)
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
app.MapPost("/api/workbooks", async (
    HttpRequest request,
    HttpContext http,
    WorkbookIngestService ingest,
    WorkbookStore store,
    IOptions<ExcelViewerOptions> opts,
    CancellationToken ct) =>
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
}).RequireRateLimiting("ingest").DisableAntiforgery();

static IResult Respond(IngestResult result, HttpContext http, WorkbookStore store, ExcelViewerOptions opts)
{
    var meta = result.Meta!;

    // PublicBaseUrl wins when set: the POST may have arrived over an internal address, but this URL
    // is about to be opened in a human's browser.
    var baseUrl = string.IsNullOrWhiteSpace(opts.PublicBaseUrl)
        ? $"{http.Request.Scheme}://{http.Request.Host}"
        : opts.PublicBaseUrl.TrimEnd('/');

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

// ── GET /api/workbooks/{hash}/original ──────────────────────────────────────────────────────
app.MapGet("/api/workbooks/{hash}/original", async (string hash, WorkbookStore store, CancellationToken ct) =>
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
});

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
