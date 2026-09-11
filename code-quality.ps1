# PowerShell script to run JetBrains CleanupCode and/or InspectCode
# Hand-authored HTML is formatted; hash-bound Graphify publication files are excluded.

param(
    [switch]$CleanupCode,
    [switch]$InspectCode,
    [string]$Profile = "Built-in: Full Cleanup",
    [string[]]$Exclude = @("*.txt", "report.xml", "Tests/TestResults/**")
)

# Set console encoding to UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# Set environment variable for .NET
$env:DOTNET_CLI_UI_LANGUAGE = "en-US"

# If neither flag is specified, run both by default
if (-not $CleanupCode -and -not $InspectCode) {
    $CleanupCode = $true
    $InspectCode = $true
}

# These files are generated as one hash-bound publication. Formatting any of them after
# graphify/build-public-map.py runs invalidates the provenance recorded in metadata.json.
$generatedPublicationExcludes = @("docs/architecture-map/**", "docs/assets/**")
$effectiveExcludes = @($Exclude + $generatedPublicationExcludes | Select-Object -Unique)
$excludeParam = $effectiveExcludes -join ";"

$exitCode = 0

# Run CleanupCode if requested
if ($CleanupCode) {
    Write-Host "=== Running JetBrains CleanupCode ===" -ForegroundColor Cyan
    Write-Host "Profile: $Profile" -ForegroundColor Gray
    Write-Host "Excluding: $excludeParam" -ForegroundColor Gray
    Write-Host ""
    
    jb cleanupcode AsposeMcpServer.sln --profile="$Profile" --exclude="$excludeParam"
    if ($LASTEXITCODE -ne 0) {
        $exitCode = $LASTEXITCODE
    }
    Write-Host ""
}

# Run InspectCode if requested
if ($InspectCode) {
    Write-Host "=== Running JetBrains InspectCode ===" -ForegroundColor Cyan
    Write-Host "Output: report.xml" -ForegroundColor Gray
    Write-Host ""
    
    jb inspectcode AsposeMcpServer.sln -o="report.xml" --exclude="$excludeParam"
    if ($LASTEXITCODE -ne 0) {
        $exitCode = $LASTEXITCODE
    }
    Write-Host ""
}

# Exit with the last non-zero exit code, or 0 if all succeeded
exit $exitCode

