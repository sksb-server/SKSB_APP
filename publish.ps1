[CmdletBinding()]
param(
    [Parameter(Position = 0, Mandatory = $false, HelpMessage = "Optional version (e.g. 1.1.0 or v1.1.0). Omit to auto-increment patch.")]
    [string]$ReleaseVersion = "",

    [Parameter(Mandatory = $false)]
    [string]$ServerHost = "http://192.168.0.20:5000",

    [Parameter(Mandatory = $false, HelpMessage = "Build and zip only; skip upload and version.json update.")]
    [switch]$SkipUpload
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$StagingDir = Join-Path $ScriptDir "publish_staging"
$ZipOut = Join-Path $ScriptDir "SKSB_Release_Package.zip"
$VersionFile = Join-Path $ScriptDir "version.json"

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " [SKSB System] Building Production Release Package" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

# 0. Fail fast if the server is offline (skip with -SkipUpload)
if (-not $SkipUpload) {
    $uri = [Uri]$ServerHost
    $reachable = $false
    try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $iar = $tcp.BeginConnect($uri.Host, $uri.Port, $null, $null)
        $reachable = $iar.AsyncWaitHandle.WaitOne(3000) -and $tcp.Connected
        $tcp.Close()
    } catch { $reachable = $false }
    if (-not $reachable) {
        throw "SERVER OFFLINE: cannot connect to $($uri.Host):$($uri.Port). Start the update service (or pass -ServerHost / -SkipUpload) and try again."
    }
}
# 1. Clean previous staging artifacts
if (Test-Path $StagingDir) { Remove-Item $StagingDir -Recurse -Force }
if (Test-Path $ZipOut) { Remove-Item $ZipOut -Force }
New-Item -ItemType Directory -Path (Join-Path $StagingDir "Templates") -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $StagingDir "AddIns") -Force | Out-Null

# 2. Determine Version: Manual Override vs Auto-Increment
$CurrentVerStr = "1.0.0"
if (Test-Path $VersionFile) {
    try {
        $json = Get-Content $VersionFile -Raw | ConvertFrom-Json
        if ($json.version) { $CurrentVerStr = $json.version }
    } catch { }
}

$CleanExplicitVer = $ReleaseVersion.Trim().TrimStart('v','V')

if (-not [string]::IsNullOrWhiteSpace($CleanExplicitVer)) {
    if (-not [System.Version]::TryParse($CleanExplicitVer, [ref]$null)) {
        Write-Error "Invalid version format '$ReleaseVersion'. Must follow standard format like 1.1.0 or 2.0.0."
        exit 1
    }
    $FinalVersion = $CleanExplicitVer
    Write-Host "Using Specified Version : v$FinalVersion" -ForegroundColor Green
} else {
    if ([System.Version]::TryParse($CurrentVerStr, [ref]$null)) {
        $v = [System.Version]::Parse($CurrentVerStr)
        $newPatch = [Math]::Max(0, $v.Build) + 1
        $FinalVersion = "$($v.Major).$($v.Minor).$newPatch"
    } else {
        $parts = $CurrentVerStr.Split('.')
        if ($parts.Length -ge 3) {
            $parts[2] = [string]([int]$parts[2] + 1)
            $FinalVersion = [string]::Join('.', $parts[0..2])
        } else {
            $FinalVersion = "1.0.1"
        }
    }
    Write-Host "Auto-Incremented Version: v$CurrentVerStr -> v$FinalVersion" -ForegroundColor Yellow
}

# 3. Build everything in Release (deterministic: never package stale/Debug binaries)
function Invoke-ReleaseBuild([string]$Target) {
    Write-Host "Building (Release): $Target" -ForegroundColor Cyan
    & dotnet build $Target -c Release --nologo -v:minimal
    if ($LASTEXITCODE -ne 0) { throw "Release build failed for '$Target' (exit code $LASTEXITCODE). Aborting publish." }
}

$SlnxPath = Join-Path $ScriptDir "SKSB_App.slnx"
$SlnxText = Get-Content $SlnxPath -Raw
Invoke-ReleaseBuild $SlnxPath
# Out-of-solution modules are optional: a failed build skips the module (its stale output is never packaged).
$SkippedModules = @()
foreach ($extra in @("ProjectManager", "ScheduleManager")) {
    $csproj = Join-Path $ScriptDir "$extra\$extra.csproj"
    if ((Test-Path $csproj) -and ($SlnxText -notmatch [regex]::Escape("$extra/$extra.csproj"))) {
        try { Invoke-ReleaseBuild $csproj }
        catch { Write-Warning "$extra skipped: $($_.Exception.Message)"; $SkippedModules += $extra }
    }
}

# 4. Collect Portal.exe and runtime dependencies (Release only)
$PortalBin = Join-Path $ScriptDir "Portal\bin\Release\net8.0-windows"
if (-not (Test-Path $PortalBin)) { throw "Portal Release output not found: $PortalBin" }
if (-not (Test-Path (Join-Path $PortalBin "Portal.exe"))) { throw "Portal.exe missing from Release output: $PortalBin" }

Write-Host "Collecting Portal binaries from: $PortalBin"
Get-ChildItem -Path $PortalBin -File | Where-Object { $_.Extension -match '\.(exe|dll|json)$' } | ForEach-Object {
    Copy-Item $_.FullName -Destination $StagingDir -Force
}
if (-not (Test-Path (Join-Path $StagingDir "Portal.exe"))) { throw "Portal.exe was not staged." }

# 5. Collect Templates (.xltx / .xltm)
$TemplatesSource = Join-Path $ScriptDir "Templates"
if (Test-Path $TemplatesSource) {
    Get-ChildItem -Path $TemplatesSource -Filter "*.xlt*" | ForEach-Object {
        Copy-Item $_.FullName -Destination (Join-Path $StagingDir "Templates") -Force
    }
}

# 6. Collect AddIns strictly from each project's Release output (never Debug/obj)
$RequiredAddInProjects = @("Core", "StaffManager", "VehicleManager")
$OptionalAddInProjects = @("ProjectManager", "ScheduleManager")
foreach ($proj in ($RequiredAddInProjects + $OptionalAddInProjects)) {
    if ($SkippedModules -contains $proj) { continue }
    $rel = Join-Path $ScriptDir "$proj\bin\Release\net8.0-windows"
    $packed = @()
    if (Test-Path (Join-Path $rel "publish")) { $packed = @(Get-ChildItem -Path (Join-Path $rel "publish") -Filter "*-packed.xll" -File) }
    if ($packed.Count -eq 0) {
        if ($RequiredAddInProjects -contains $proj) { throw "No Release *-packed.xll found for $proj in $rel\publish" }
        continue
    }
    foreach ($f in $packed) {
        Copy-Item $f.FullName -Destination (Join-Path $StagingDir "AddIns\$($f.Name.Replace('-packed',''))") -Force
    }
    Get-ChildItem -Path $rel -Filter "*.dna" -File | ForEach-Object {
        Copy-Item $_.FullName -Destination (Join-Path $StagingDir "AddIns") -Force
    }
}

# 7. Write and stage version.json (repo copy is written only after a successful upload)
$Manifest = @{
    version = $FinalVersion
    lastUpdatedUtc = (Get-Date).ToUniversalTime().ToString("o")
    environment = "Production"
}
$ManifestJson = $Manifest | ConvertTo-Json -Depth 2
$ManifestJson | Out-File (Join-Path $StagingDir "version.json") -Encoding utf8
# 8. Compress Package
Compress-Archive -Path "$StagingDir\*" -DestinationPath $ZipOut -Force
Remove-Item $StagingDir -Recurse -Force
Write-Host "Package bundled: $ZipOut" -ForegroundColor Green

# 9. Upload to server (any failure aborts with a non-zero exit code)
if ($SkipUpload) { Write-Host "SkipUpload set: package built at $ZipOut (not uploaded)." -ForegroundColor Yellow; return }
Write-Host "Deploying to server at: $ServerHost/api/update/publish..." -ForegroundColor Yellow
$PublishUrl = "$ServerHost/api/update/publish"
$curlCmd = Get-Command curl.exe -ErrorAction SilentlyContinue
if ($curlCmd -and $curlCmd.Source) {
    & $curlCmd.Source -X POST $PublishUrl -F "package=@$ZipOut" -F "version=$FinalVersion" --fail --silent --show-error | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "curl upload failed with exit code $LASTEXITCODE" }
} else {
    Add-Type -AssemblyName System.Net.Http
    $client = New-Object System.Net.Http.HttpClient
    $client.Timeout = [TimeSpan]::FromMinutes(10)
    $content = New-Object System.Net.Http.MultipartFormDataContent
    $fileBytes = [System.IO.File]::ReadAllBytes($ZipOut)
    $zipContent = New-Object System.Net.Http.ByteArrayContent(,$fileBytes)
    $zipContent.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse("application/zip")
    $content.Add($zipContent, "package", "SKSB_Update.zip")
    $content.Add((New-Object System.Net.Http.StringContent($FinalVersion)), "version")
    $resp = $client.PostAsync($PublishUrl, $content).GetAwaiter().GetResult()
    if (-not $resp.IsSuccessStatusCode) { throw "Server upload returned status: $($resp.StatusCode)" }
}

$ManifestJson | Out-File $VersionFile -Encoding utf8
Write-Host "DEPLOYMENT SUCCESS! Server activated version: v$FinalVersion" -ForegroundColor Green