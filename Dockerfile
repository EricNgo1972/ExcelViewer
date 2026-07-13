# syntax=docker/dockerfile:1

# ── Build ────────────────────────────────────────────────────────────────────
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src
COPY Directory.Build.props ./
COPY MK.ExcelViewer/MK.ExcelViewer.csproj MK.ExcelViewer/
RUN dotnet restore MK.ExcelViewer/MK.ExcelViewer.csproj
COPY MK.ExcelViewer/ MK.ExcelViewer/
RUN dotnet publish MK.ExcelViewer/MK.ExcelViewer.csproj -c Release -o /app --no-restore

# ── Runtime ──────────────────────────────────────────────────────────────────
# Pure managed: ClosedXML parses in-process. No LibreOffice, no Poppler, no fonts —
# this image is the ASP.NET runtime plus a few MB of app.
FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS runtime
WORKDIR /app

# The aspnet:8.0 base is Debian slim and ships NEITHER curl NOR wget. Without this the HEALTHCHECK
# below can never pass, and the container runs forever marked unhealthy — it boots fine, so the
# failure looks like a mystery rather than a missing binary. (mk-FileConverter only gets away with
# `wget` because its LibreOffice apt install happens to pull it in; we install nothing else.)
RUN apt-get update \
 && apt-get install -y --no-install-recommends curl \
 && rm -rf /var/lib/apt/lists/*

COPY --from=build /app ./

# Durable state — the content-addressed workbook store — is written under the hard-coded Linux path
# /var/lib/mk-excelviewer (see the state-root logic in Program.cs). Symlink that onto /data so it
# lives on the mounted volume and survives container recreation. Same trick as spc-docconverter
# and MemberList.
RUN mkdir -p /data && ln -s /data /var/lib/mk-excelviewer

ENV ASPNETCORE_HTTP_PORTS=8080
EXPOSE 8080

# The cheap liveness probe, not /api/health — we have no engine to prove alive, so there's nothing
# for the deeper check to earn here.
HEALTHCHECK --interval=30s --timeout=10s --start-period=15s --retries=3 \
    CMD ["sh", "-c", "curl -fsS http://localhost:8080/health >/dev/null || exit 1"]

ENTRYPOINT ["dotnet", "MK.ExcelViewer.dll"]
