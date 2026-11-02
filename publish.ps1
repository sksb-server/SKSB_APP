[CmdletBinding()]
param(
    [Parameter(Position = 0, Mandatory = $false, HelpMessage = "Optional version (e.g. 1.1.0 or v1.1.0). Omit to auto-increment patch.")]
    [string]$ReleaseVersion = "",

    [Parameter(Mandatory = $false)]
    [string]$ServerHost = "http://192.168.0.20:5000"
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$StagingDir = Join-Path $ScriptDir "publish_staging"
$ZipOut = Join-Path $ScriptDir "SKSB_Release_Package.zip"
$VersionFile = Join-Path $ScriptDir "version.json"

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " [SKSB System] Building Production Release Package" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

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

# 3. Collect Portal.exe and runtime dependencies
$PortalBin = Join-Path $ScriptDir "Portal\bin\Release\net8.0-windows"
if (-not (Test-Path $PortalBin)) {
    $PortalBin = Join-Path $ScriptDir "Portal\bin\Debug\net8.0-windows"
}
if (-not (Test-Path $PortalBin)) {
    Write-Error "Could not locate Portal build artifacts. Please build Portal project first."
    exit 1
}

Write-Host "Collecting Portal binaries from: $PortalBin"
Get-ChildItem -Path $PortalBin -File | Where-Object { $_.Extension -match '\.(exe|dll|json|runtimeconfig\.json)$' } | ForEach-Object {
    Copy-Item $_.FullName -Destination $StagingDir -Force
}

# 4. Collect Templates (.xltx / .xltm)
$TemplatesSource = Join-Path $ScriptDir "Templates"
if (Test-Path $TemplatesSource) {
    Get-ChildItem -Path $TemplatesSource -Filter "*.xlt*" | ForEach-Object {
        Copy-Item $_.FullName -Destination (Join-Path $StagingDir "Templates") -Force
    }
}

# 5. Collect AddIns (.xll / .dna)
Get-ChildItem -Path $ScriptDir -Recurse -Filter "*-packed.xll" | ForEach-Object {
    Copy-Item $_.FullName -Destination (Join-Path $StagingDir "AddIns\$($_.Name.Replace('-packed',''))") -Force
}
Get-ChildItem -Path $ScriptDir -Recurse -Filter "*.dna" | Where-Object { $_.FullName -notmatch "publish_staging|obj|bin" } | ForEach-Object {
    Copy-Item $_.FullName -Destination (Join-Path $StagingDir "AddIns") -Force
}

# 6. Write and Stage version.json
$Manifest = @{
    version = $FinalVersion
    lastUpdatedUtc = (Get-Date).ToUniversalTime().ToString("o")
    environment = "Production"
}
$ManifestJson = $Manifest | ConvertTo-Json -Depth 2
$ManifestJson | Out-File (Join-Path $StagingDir "version.json") -Encoding utf8
$ManifestJson | Out-File $VersionFile -Encoding utf8

# 7. Compress Package
Compress-Archive -Path "$StagingDir\*" -DestinationPath $ZipOut -Force
Remove-Item $StagingDir -Recurse -Force
Write-Host "Package bundled: $ZipOut" -ForegroundColor Green

# 8. Upload to Linux Server (Universal curl command execution)
try {
    Write-Host "Deploying to server at: $ServerHost/api/update/publish..." -ForegroundColor Yellow
    $PublishUrl = "$ServerHost/api/update/publish"
    
    # Locate curl safely without using PowerShell 7 ?. operator
    $curlCmd = Get-Command curl.exe -ErrorAction SilentlyContinue
    if ($curlCmd -and $curlCmd.Source) {
        $curlPath = $curlCmd.Source
        & $curlPath -X POST $PublishUrl `
                    -F "package=@$ZipOut" `
                    -F "version=$FinalVersion" `
                    --fail --silent --show-error | Out-Null
        
        Write-Host "DEPLOYMENT SUCCESS! Server activated version: v$FinalVersion" -ForegroundColor Green
    } else {
        # Fallback using System.Net.Http for environments without curl.exe
        [System.Reflection.Assembly]::LoadWithPartialName("System.Net.Http") | Out-Null
        $fileBytes = [System.IO.File]::ReadAllBytes($ZipOut)
        $client = New-Object System.Net.Http.HttpClient
        $content = New-Object System.Net.Http.MultipartFormDataContent
        
        $zipContent = New-Object System.Net.Http.ByteArrayContent($fileBytes, 0, $fileBytes.Length)
        $zipContent.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse("application/zip")
        
        $content.Add($zipContent, "package", "SKSB_Update.zip")
        $content.Add((New-Object System.Net.Http.StringContent($FinalVersion)), "version")

        $resp = $client.PostAsync($PublishUrl, $content).GetAwaiter().GetResult()
        if ($resp.IsSuccessStatusCode) {
            Write-Host "DEPLOYMENT SUCCESS! Server activated version: v$FinalVersion" -ForegroundColor Green
        } else {
            Write-Warning "Server upload returned status: $($resp.StatusCode)"
        }
    }
} catch {
    Write-Warning "Deployment failed: $_"
}