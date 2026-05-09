# =============================================================================
#  Navisworks MCP Bridge — installer
# =============================================================================
#  Copies the FULL build output to the Navisworks plugin folder.
#
#  Why this matters: when targeting .NET Framework 4.8, System.Text.Json is
#  pulled from NuGet and brings ~6 transitive dependencies (System.Memory,
#  System.Buffers, System.Numerics.Vectors, etc.).  All of them must sit
#  next to NavisworksMcpAddin.dll or .NET will silently fail to load it
#  and the ribbon button never appears.
# =============================================================================

param(
    [string]$Year = "2027",
    [string]$BuildDir = "bin\Release"
)

$ErrorActionPreference = "Stop"

$source = Resolve-Path $BuildDir
$dest   = "$env:APPDATA\Autodesk\Navisworks Manage $Year\Plugins\MCPBridge"

Write-Host "Source: $source"
Write-Host "Dest  : $dest"
Write-Host ""

# Sanity check: confirm the build actually produced a DLL.
$builtDll = Get-ChildItem $source -Filter "MCPBridge.dll" -ErrorAction SilentlyContinue
if (-not $builtDll) {
    Write-Host "ERROR: MCPBridge.dll is not in $source." -ForegroundColor Red
    Write-Host "       The build produced nothing — check the dotnet build output for errors." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Files actually in $source :" -ForegroundColor Yellow
    Get-ChildItem $source -ErrorAction SilentlyContinue | ForEach-Object { Write-Host "  $($_.Name)" }
    exit 1
}

# Wipe the destination first so we don't keep stale dependencies around
if (Test-Path $dest) {
    Write-Host "Cleaning existing folder..." -ForegroundColor Yellow
    Remove-Item "$dest\*" -Force -Recurse
}
New-Item -ItemType Directory -Force $dest | Out-Null

# Copy every DLL, every PDB, the manifest, and any runtimes\ subfolders
Write-Host "Copying files..." -ForegroundColor Cyan
$copied = @()
Get-ChildItem $source -File | ForEach-Object {
    if ($_.Extension -in ".dll", ".addin", ".pdb", ".xml", ".config") {
        Copy-Item $_.FullName $dest -Force
        $copied += $_.Name
    }
}

# Some NuGet packages drop platform-specific DLLs into a runtimes\ folder.
# Copy that too so the platform-correct binaries are available.
$runtimes = Join-Path $source "runtimes"
if (Test-Path $runtimes) {
    Write-Host "Copying runtimes\ subfolder..." -ForegroundColor Cyan
    Copy-Item $runtimes $dest -Recurse -Force
}

# Unblock everything (Mark-of-the-Web prevents .NET from loading downloaded DLLs)
Write-Host "Unblocking files..." -ForegroundColor Cyan
Get-ChildItem $dest -Recurse -File | Unblock-File -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "Installed files:" -ForegroundColor Green
Get-ChildItem $dest -Recurse -File | ForEach-Object {
    $rel = $_.FullName.Substring($dest.Length + 1)
    Write-Host ("  {0,-50} {1,10:N0} bytes" -f $rel, $_.Length)
}

Write-Host ""
Write-Host "Done. Restart Navisworks $Year to load the addin." -ForegroundColor Green
