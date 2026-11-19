[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [ValidateSet('Portal', 'Schedule', 'ScheduleManager', 'Staff', 'StaffManager', 'Project', 'ProjectManager', 'Vehicle', 'VehicleManager')]
    [string]$Module = 'Portal',
    [switch]$Kill,
    [switch]$NoBuild,
    [switch]$OpenVS,
    [switch]$OpenVSCode,
    [switch]$Wait,
    [string]$Worktree
)

$ErrorActionPreference = 'Stop'

$worktreeRoot = $PSScriptRoot
if ($Worktree) {
    if (Test-Path $Worktree) {
        $worktreeRoot = (Resolve-Path $Worktree).Path
    } else {
        $entries = @()
        $cur = $null
        foreach ($line in (git -C $PSScriptRoot worktree list --porcelain)) {
            if ($line -like 'worktree *') {
                if ($cur) { $entries += $cur }
                $path = $line.Substring(9).Trim()
                $cur = [pscustomobject]@{ Path = $path; Leaf = (Split-Path $path -Leaf); Branch = '' }
            } elseif ($cur -and $line -like 'branch *') {
                $cur.Branch = ($line.Substring(7).Trim() -replace '^refs/heads/', '')
            }
        }
        if ($cur) { $entries += $cur }

        $match = @($entries | Where-Object { $_.Leaf -ieq $Worktree -or $_.Branch -ieq $Worktree }) | Select-Object -First 1
        if (-not $match) {
            $match = @($entries | Where-Object { $_.Leaf -like "*$Worktree*" -or $_.Branch -like "*$Worktree*" }) | Select-Object -First 1
        }
        if (-not $match) {
            $list = ($entries | ForEach-Object { "  $($_.Leaf)  [$($_.Branch)]  $($_.Path)" }) -join "`n"
            throw "Worktree '$Worktree' not found. Available worktrees:`n$list"
        }
        $worktreeRoot = $match.Path
    }
}
$worktreeRoot = [System.IO.Path]::GetFullPath($worktreeRoot)
$env:SKSB_SOLUTION_ROOT = $worktreeRoot

$branch = ''
try { $branch = (git -C $worktreeRoot rev-parse --abbrev-ref HEAD 2>$null) } catch { }
Write-Host "Worktree: $(Split-Path $worktreeRoot -Leaf)" -ForegroundColor Cyan
Write-Host "Branch:   $branch" -ForegroundColor Cyan

if ($OpenVS) {
    $vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
    $devenv = $null
    if (Test-Path $vswhere) {
        $devenv = & $vswhere -latest -prerelease -products * -requires Microsoft.Component.MSBuild -property productPath | Select-Object -First 1
        if (-not $devenv) {
            $devenv = & $vswhere -latest -prerelease -products * -property productPath | Select-Object -First 1
        }
    }
    if (-not $devenv -or -not (Test-Path $devenv)) {
        $devenv = 'C:\Program Files\Microsoft Visual Studio\18\Community\Common7\IDE\devenv.exe'
    }
    if (-not (Test-Path $devenv)) { throw 'devenv.exe not found (vswhere and default Visual Studio 18 Community path checked).' }

    $slnPath = "$worktreeRoot\SKSB_App.slnx"
    if (-not (Test-Path $slnPath)) { throw "Solution not found: $slnPath" }

    Write-Host "Opened Visual Studio: $slnPath" -ForegroundColor Green
    Start-Process -FilePath $devenv -ArgumentList "`"$slnPath`"" -WorkingDirectory $worktreeRoot -WindowStyle Normal -Wait:$Wait
    return
}

if ($OpenVSCode) {
    $codeCmd = Get-Command code.cmd -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $codeCmd) { $codeCmd = Get-Command code -ErrorAction SilentlyContinue | Select-Object -First 1 }
    if (-not $codeCmd) { throw "VS Code 'code' command not found on PATH." }

    Write-Host "Opened VS Code: $worktreeRoot" -ForegroundColor Green
    $codeArgs = if ($Wait) { '-n', '-w', "`"$worktreeRoot`"" } else { '-n', "`"$worktreeRoot`"" }
    Start-Process -FilePath $codeCmd.Source -ArgumentList $codeArgs -WorkingDirectory $worktreeRoot -WindowStyle Normal -Wait:$Wait
    return
}

if ($Module -eq 'Portal') {
    if ($Kill) {
        Get-Process Portal, EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
    }

    if (-not $NoBuild) {
        dotnet build "$worktreeRoot\SKSB_App.slnx" -c Debug
        if ($LASTEXITCODE -ne 0) { throw "Solution build failed (exit $LASTEXITCODE)." }
    }

    if (-not $Kill -and (Get-Process Portal -ErrorAction SilentlyContinue)) {
        Write-Host "WARNING: Portal is already running. Portal is single-instance and will focus the existing Portal; use -Kill to restart it from this worktree." -ForegroundColor Yellow
    }

    $portalDir = "$worktreeRoot\Portal\bin\Debug\net8.0-windows"
    $portalExe = "$portalDir\Portal.exe"
    if (-not (Test-Path $portalExe)) { throw "Portal.exe not found: $portalExe" }

    Start-Process -FilePath $portalExe -WorkingDirectory $portalDir -WindowStyle Normal
    return
}

# ---- Direct module mode ----
switch -Regex ($Module) {
    '^Schedule' { $projectName = 'ScheduleManager'; $templateName = 'Schedule_Manager_Template.xltx' }
    '^Staff'    { $projectName = 'StaffManager';    $templateName = 'Staff_Manager_Template.xltx' }
    '^Project'  { $projectName = 'ProjectManager';  $templateName = 'Project_Manager_Template.xltx' }
    '^Vehicle'  { $projectName = 'VehicleManager';  $templateName = 'Vehicle_Manager_Template.xltx' }
}

if ($Kill) {
    Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
}

if (-not $NoBuild) {
    if ($projectName -in 'ScheduleManager', 'StaffManager' -and
        -not (Test-Path "$worktreeRoot\ZkBridge\bin\Debug\net8.0-windows\ZkBridge.exe")) {
        dotnet build "$worktreeRoot\ZkBridge\ZkBridge.csproj" -c Debug
        if ($LASTEXITCODE -ne 0) { throw "ZkBridge build failed (exit $LASTEXITCODE)." }
    }
    dotnet build "$worktreeRoot\$projectName\$projectName.csproj" -c Debug
    if ($LASTEXITCODE -ne 0) { throw "$projectName build failed (exit $LASTEXITCODE)." }
}

# ---- Locate Excel ----
$excelPath = $null
foreach ($key in 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe',
                 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\excel.exe') {
    try {
        $p = (Get-ItemProperty -Path $key -ErrorAction Stop).'(default)'
        if (-not $p) { $p = (Get-Item $key).GetValue('') }
        if ($p -and (Test-Path $p)) { $excelPath = $p; break }
    } catch { }
}
if (-not $excelPath) {
    foreach ($c in "$env:ProgramFiles\Microsoft Office\root\Office16\EXCEL.EXE",
                   "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\EXCEL.EXE") {
        if ($c -and (Test-Path $c)) { $excelPath = $c; break }
    }
}
if (-not $excelPath) { throw 'excel.exe not found (registry App Paths and standard Office16 locations checked).' }

# ---- Bitness via PE header ----
$fs = [System.IO.File]::OpenRead($excelPath)
try {
    $br = New-Object System.IO.BinaryReader($fs)
    $fs.Seek(0x3C, 'Begin') | Out-Null
    $peOffset = $br.ReadInt32()
    $fs.Seek($peOffset + 4, 'Begin') | Out-Null
    $machine = $br.ReadUInt16()
} finally { $fs.Dispose() }
$suffix = if ($machine -eq 0x8664) { '-AddIn64.xll' } else { '-AddIn.xll' }

# ---- Stage shadow copy ----
$shadowRoot = Join-Path $env:TEMP 'SKSB_Debug_Shadow'
New-Item -ItemType Directory -Force -Path $shadowRoot | Out-Null
Get-ChildItem $shadowRoot -Directory -ErrorAction SilentlyContinue | ForEach-Object {
    try { Remove-Item $_.FullName -Recurse -Force -ErrorAction Stop } catch { }
}

$sessionDir = Join-Path $shadowRoot "${projectName}_$(Get-Date -Format 'yyyyMMdd_HHmmss_fff')_$([Guid]::NewGuid().ToString('N').Substring(0,6))"
New-Item -ItemType Directory -Force -Path $sessionDir | Out-Null

foreach ($src in "$worktreeRoot\Core\bin\Debug\net8.0-windows", "$worktreeRoot\$projectName\bin\Debug\net8.0-windows") {
    if (-not (Test-Path $src)) { throw "Build output not found: $src" }
    Get-ChildItem $src -File | Copy-Item -Destination $sessionDir -Force
}

Set-Content -Path (Join-Path $sessionDir '.sksb_solution_root') -Value $worktreeRoot -NoNewline

# Best-effort trusted locations
try {
    $tlRoot = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations'
    foreach ($loc in @{ Name = 'SKSB_App_DebugShadow'; Path = $shadowRoot }, @{ Name = 'SKSB_App_DebugRepo'; Path = $worktreeRoot }) {
        $k = Join-Path $tlRoot $loc.Name
        if (-not (Test-Path $k)) { New-Item -Path $k -Force | Out-Null }
        New-ItemProperty -Path $k -Name 'Path' -Value $loc.Path -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $k -Name 'AllowSubfolders' -Value 1 -PropertyType DWord -Force | Out-Null
    }
} catch {
    Write-Host "Could not register trusted locations: $($_.Exception.Message)" -ForegroundColor Yellow
}

$templatePath  = "$worktreeRoot\Templates\$templateName"
$shadowCoreXll = Join-Path $sessionDir "Core$suffix"
$shadowModuleXll = Join-Path $sessionDir "$projectName$suffix"

foreach ($f in $templatePath, $shadowCoreXll, $shadowModuleXll) {
    if (-not (Test-Path $f)) { throw "Required file not found: $f" }
}

Start-Process -FilePath $excelPath -ArgumentList "/x /t `"$templatePath`" `"$shadowCoreXll`" `"$shadowModuleXll`"" -WindowStyle Normal
