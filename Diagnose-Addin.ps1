# =============================================================================
#  Navisworks MCP Bridge — addin load diagnostic
# =============================================================================
#  Run this in PowerShell from anywhere — it does NOT require Navisworks
#  to be running.  It checks the five most common reasons a Navisworks
#  plugin fails to appear in the Add-Ins ribbon:
#
#    1. Wrong target folder (e.g. you copied to "Plugins" instead of
#       "Plugins\MCPBridge\")
#    2. Missing .addin manifest, or manifest not in the same folder as DLL
#    3. Folder name does not match the Plugin attribute name
#    4. .NET runtime mismatch (built for 8.0, but Navisworks loads 4.x)
#    5. DLL is unblocked from internet zone (the Mark-of-the-Web problem)
# =============================================================================

$ErrorActionPreference = "Continue"

Write-Host "=== Navisworks MCP Bridge diagnostic ===" -ForegroundColor Cyan
Write-Host ""

# ── 1. Find every Navisworks Plugins folder on this user's profile ─────────
Write-Host "[1] Scanning %APPDATA% for Navisworks Plugins folders..."
$pluginRoots = Get-ChildItem "$env:APPDATA\Autodesk" -Directory -ErrorAction SilentlyContinue |
               Where-Object { $_.Name -match "Navisworks (Manage|Simulate)" } |
               ForEach-Object { Join-Path $_.FullName "Plugins" } |
               Where-Object { Test-Path $_ }

if (-not $pluginRoots) {
    Write-Host "  NONE FOUND. Open Navisworks once to create the profile, then re-run." -ForegroundColor Red
    exit 1
}
$pluginRoots | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }
Write-Host ""

# ── 2. Look for the MCPBridge subfolder in each ────────────────────────────
Write-Host "[2] Looking for MCPBridge subfolder..."
$found = @()
foreach ($root in $pluginRoots) {
    $mcp = Join-Path $root "MCPBridge"
    if (Test-Path $mcp) {
        $found += $mcp
        Write-Host "  FOUND: $mcp" -ForegroundColor Green
    } else {
        Write-Host "  NOT in: $root" -ForegroundColor Yellow
    }
}
if (-not $found) {
    Write-Host ""
    Write-Host "  PROBLEM: No MCPBridge subfolder. Create it and copy the build output." -ForegroundColor Red
    Write-Host "  Run:" -ForegroundColor Yellow
    foreach ($root in $pluginRoots) {
        Write-Host "    New-Item -ItemType Directory -Force '$root\MCPBridge'"
    }
    exit 1
}
Write-Host ""

# ── 3. Per-folder content + manifest validation ─────────────────────────────
Write-Host "[3] Validating each install..."
foreach ($mcp in $found) {
    Write-Host ""
    Write-Host "--- $mcp ---" -ForegroundColor Cyan
    $files = Get-ChildItem $mcp -ErrorAction SilentlyContinue
    if (-not $files) { Write-Host "  EMPTY FOLDER." -ForegroundColor Red; continue }
    $files | ForEach-Object { Write-Host ("  {0,-40} {1,10} bytes" -f $_.Name, $_.Length) }

    $dll      = Get-Item (Join-Path $mcp "NavisworksMcpAddin.dll") -ErrorAction SilentlyContinue
    $manifest = Get-Item (Join-Path $mcp "MCPBridge.addin")        -ErrorAction SilentlyContinue

    if (-not $dll) {
        Write-Host "  PROBLEM: NavisworksMcpAddin.dll is missing." -ForegroundColor Red
    }
    if (-not $manifest) {
        Write-Host "  PROBLEM: MCPBridge.addin manifest is missing." -ForegroundColor Red
    }

    # ── 3a. .addin manifest sanity check ──────────────────────────
    if ($manifest) {
        try {
            [xml]$xml = Get-Content $manifest.FullName -ErrorAction Stop
            $plug = $xml.RibbonInfo.Plugins.Plugin
            $asm  = $plug.Assembly
            Write-Host "  manifest Plugin Name = $($plug.Name)" -ForegroundColor Gray
            Write-Host "  manifest Assembly    = $asm" -ForegroundColor Gray
            if ($plug.Name -ne (Split-Path $mcp -Leaf)) {
                Write-Host "  PROBLEM: Plugin Name '$($plug.Name)' does NOT match folder '$(Split-Path $mcp -Leaf)'." -ForegroundColor Red
                Write-Host "           Navisworks rejects mismatched names. Rename the folder OR edit the manifest." -ForegroundColor Yellow
            }
            $expectedDll = Join-Path $mcp $asm
            if (-not (Test-Path $expectedDll)) {
                Write-Host "  PROBLEM: Assembly '$asm' referenced by manifest is NOT in the folder." -ForegroundColor Red
            }
        } catch {
            Write-Host "  PROBLEM: manifest XML is invalid — $_" -ForegroundColor Red
        }
    }

    # ── 3b. Mark-of-the-Web (downloaded from internet) check ─────
    if ($dll) {
        $zone = Get-Item $dll.FullName -Stream Zone.Identifier -ErrorAction SilentlyContinue
        if ($zone) {
            Write-Host "  PROBLEM: DLL has Mark-of-the-Web → Windows blocks .NET from loading it." -ForegroundColor Red
            Write-Host "           Fix:  Unblock-File '$($dll.FullName)'" -ForegroundColor Yellow
        }
    }

    # ── 3c. Check DLL target framework ────────────────────────────
    if ($dll) {
        try {
            $asm = [System.Reflection.Assembly]::ReflectionOnlyLoadFrom($dll.FullName)
            $tfm = $asm.GetCustomAttributesData() |
                   Where-Object { $_.AttributeType.Name -eq "TargetFrameworkAttribute" } |
                   ForEach-Object { $_.ConstructorArguments[0].Value }
            Write-Host "  DLL TargetFramework  = $tfm" -ForegroundColor Gray
            if ($tfm -match "net8\.0|net7\.0|net6\.0|net5\.0|netcoreapp") {
                Write-Host "  PROBLEM: Built for modern .NET — Navisworks 2025/2026/2027 only loads .NET Framework 4.x." -ForegroundColor Red
                Write-Host "           Fix: change <TargetFramework> in the .csproj to net48, then rebuild." -ForegroundColor Yellow
            }
        } catch {
            Write-Host "  Could not inspect DLL framework: $_" -ForegroundColor Yellow
        }
    }
}

# ── 4. Show the addin's own startup log if it exists ───────────────────────
Write-Host ""
Write-Host "[4] Addin startup log (if present)..."
$log = Join-Path $env:TEMP "navisworks_mcp_addin.log"
if (Test-Path $log) {
    Write-Host "  Reading: $log" -ForegroundColor Gray
    Get-Content $log -Tail 30
} else {
    Write-Host "  No log file at $log — DLL likely never loaded at all." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=== Diagnostic complete ===" -ForegroundColor Cyan
