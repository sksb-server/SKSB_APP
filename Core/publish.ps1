param (
    [string]$Version = "1.0.1",
    [string]$ServerHost = "http://192.168.0.20:5000",
    [string]$Configuration = "Release",
    [string]$TargetFramework = "net8.0-windows"
)

$ErrorActionPreference = "Stop"

Write-Host ">>> 1. Compiling Release Excel Add-in (v$Version)..." -ForegroundColor Cyan

# Locate project file accurately
$projectDir = $PSScriptRoot
$csprojFile = Get-ChildItem -Path $projectDir -Filter "*.csproj" -Recurse | Select-Object -First 1

if (-not $csprojFile) {
    throw "Could not locate a .csproj file in $projectDir or its subdirectories."
}

$buildOutputDir = Join-Path $csprojFile.DirectoryName "bin\$Configuration\$TargetFramework"
$distDir = Join-Path $projectDir "publish_dist"
$stagingDir = Join-Path $projectDir "staging"

# Clean target directories safely
if (Test-Path $distDir) { 
    Remove-Item $distDir -Recurse -Force 
}
if (Test-Path $stagingDir) { 
    Remove-Item $stagingDir -Recurse -Force 
}

New-Item -Path $distDir -ItemType Directory -Force | Out-Null
New-Item -Path $stagingDir -ItemType Directory -Force | Out-Null

# Compile using clean array arguments (avoids backtick line-continuation bugs)
$buildArgs = @(
    "build",
    $csprojFile.FullName,
    "-c", $Configuration,
    "-p:AssemblyVersion=$Version",
    "-p:FileVersion=$Version",
    "-p:Version=$Version"
)

& dotnet $buildArgs

if ($LASTEXITCODE -ne 0) {
    throw "Build failed with exit code $LASTEXITCODE. Check compile errors above."
}

Write-Host ">>> 2. Packaging distribution archive..." -ForegroundColor Cyan

# Files required for 32-bit and 64-bit client execution
$requiredFiles = @(
    "DailyScheduler-AddIn.xll",
    "DailyScheduler-AddIn64.xll",
    "DailyScheduler.dll",
    "DailyScheduler.runtimeconfig.json",
    "DailyScheduler.deps.json"
)

foreach ($file in $requiredFiles) {
    $sourcePath = Join-Path $buildOutputDir $file
    if (Test-Path $sourcePath) {
        Copy-Item -Path $sourcePath -Destination $stagingDir -Force
    } else {
        Write-Warning "File not found in build output: $sourcePath"
    }
}

# Add launcher script into the package
$launcherBatContent = "@echo off`r`n" +
"setlocal`r`n" +
"reg query ""HKLM\Software\Microsoft\Office\ClickToRun\Configuration"" /v Platform 2>nul | findstr /I ""x64"" >nul`r`n" +
"if %errorlevel% equ 0 (`r`n" +
"    start """" ""%~dp0DailyScheduler-AddIn64.xll""`r`n" +
") else (`r`n" +
"    start """" ""%~dp0DailyScheduler-AddIn.xll""`r`n" +
")`r`n" +
"exit"

Set-Content -Path (Join-Path $stagingDir "Launch-Allocation.bat") -Value $launcherBatContent -Encoding ASCII

# Zip the package safely using .NET compression to avoid PowerShell Compress-Archive bugs
Add-Type -AssemblyName System.IO.Compression.FileSystem
$packageZipPath = Join-Path $distDir "DailyScheduler.zip"

if (Test-Path $packageZipPath) {
    Remove-Item $packageZipPath -Force
}

[System.IO.Compression.ZipFile]::CreateFromDirectory($stagingDir, $packageZipPath)

$packageFile = Get-Item $packageZipPath
Write-Host "Packaged distribution: $($packageFile.Name) ($([math]::Round($packageFile.Length / 1MB, 2)) MB)" -ForegroundColor Gray

Write-Host ">>> 3. Generating version.json..." -ForegroundColor Cyan
$manifest = @{
    Version = $Version
    PackageName = $packageFile.Name
    ReleaseNotes = "Automated release v$Version"
} | ConvertTo-Json -Compress

$manifestPath = Join-Path $distDir "version.json"
Set-Content -Path $manifestPath -Value $manifest -Encoding UTF8

Write-Host ">>> 4. Deploying package and manifest to server..." -ForegroundColor Green
$uri = "$ServerHost/api/update/publish"

Add-Type -AssemblyName System.Net.Http
$httpClient = [System.Net.Http.HttpClient]::new()
$httpClient.Timeout = [System.TimeSpan]::FromMinutes(5)
$multipartContent = [System.Net.Http.MultipartFormDataContent]::new()

$fileStream = $null
try {
    # 1. Attach version.json
    $versionFileBytes = [System.IO.File]::ReadAllBytes($manifestPath)
    $versionContent = [System.Net.Http.ByteArrayContent]::new($versionFileBytes)
    $multipartContent.Add($versionContent, "version", [System.IO.Path]::GetFileName($manifestPath))

    # 2. Attach package stream
    $fileStream = [System.IO.File]::OpenRead($packageFile.FullName)
    $streamContent = [System.Net.Http.StreamContent]::new($fileStream)
    $multipartContent.Add($streamContent, "package", $packageFile.Name)

    $response = $httpClient.PostAsync($uri, $multipartContent).GetAwaiter().GetResult()
    
    if (-not $response.IsSuccessStatusCode) {
        $errorBody = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
        throw "Server returned error: $($response.StatusCode) - $errorBody"
    }

    Write-Host ">>> DEPLOYMENT COMPLETE (v$Version) <<<" -ForegroundColor Green
}
finally {
    if ($null -ne $fileStream) { $fileStream.Dispose() }
    $multipartContent.Dispose()
    $httpClient.Dispose()
    if (Test-Path $stagingDir) { Remove-Item $stagingDir -Recurse -Force }
}