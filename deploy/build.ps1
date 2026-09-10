# PowerShell script to build with UTF-8 encoding
# This script sets UTF-8 encoding to prevent Chinese character display issues
#
# Usage:
#   .\build.ps1          # Build Release (default)
#   .\build.ps1 -Debug   # Build Debug

param(
    [switch]$Debug
)

# Set console encoding to UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# Set environment variable for .NET
$env:DOTNET_CLI_UI_LANGUAGE = "en-US"

# Determine configuration
$configuration = if ($Debug) { "Debug" } else { "Release" }

# Build the main project explicitly
# --no-incremental, because a build that recompiles nothing reports no warnings, and a warning
# count from such a build reads exactly like a clean one. Successive incremental runs of this
# script printed 2, then 4, then 0 warnings for the same tree, and the zero was the artefact
# (R18-BUILD01).
$output = & dotnet build AsposeMcpServer.csproj --configuration $configuration --no-incremental 2>&1
$buildExitCode = $LASTEXITCODE
$output | ForEach-Object { Write-Host $_ }

if ($buildExitCode -ne 0) { exit $buildExitCode }

# The project's rule is 0 warnings and 0 errors. `dotnet build` exits 0 with warnings, so the rule
# was left to whoever read the output rather than being enforced by the thing that checks it.
$warnings = @($output | Select-String -Pattern ": warning [A-Z]+[0-9]+" -AllMatches)
if ($warnings.Count -gt 0) {
    Write-Host ""
    Write-Host "=== Build Warnings ===" -ForegroundColor Red
    Write-Host "$($warnings.Count) warning line(s). This project requires 0 warnings and 0 errors." `
        -ForegroundColor Red
    exit 1
}

exit 0

