# PowerShell script to run unit tests with UTF-8 encoding
# This script sets UTF-8 encoding to prevent Chinese character display issues

param(
    [switch]$Verbose,
    [switch]$NoBuild,
    [switch]$Coverage,
    [string]$Filter,
    [switch]$SkipLicense  # Skip license loading, force evaluation mode
)

# Set console encoding to UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# Set environment variable for .NET
$env:DOTNET_CLI_UI_LANGUAGE = "en-US"

# Set license skip flag (if SkipLicense parameter is specified)
if ($SkipLicense) {
    $env:SKIP_ASPOSE_LICENSE = "true"
    Write-Host "=== Running Tests in Evaluation Mode (License Skipped) ===" -ForegroundColor Yellow
} else {
    # Clear environment variable (if previously set)
    Remove-Item Env:\SKIP_ASPOSE_LICENSE -ErrorAction SilentlyContinue
    Write-Host "=== Running Unit Tests ===" -ForegroundColor Cyan
}
Write-Host ""

# Build test arguments
$testArgs = @()

if ($Verbose) {
    $testArgs += "--verbosity", "normal"
} else {
    $testArgs += "--verbosity", "minimal"
}

if ($NoBuild) {
    $testArgs += "--no-build"
}

if ($Coverage) {
    $testArgs += "--collect:`"XPlat Code Coverage`""
}

if ($Filter) {
    $testArgs += "--filter", $Filter
}

# Add logger for test results
$testArgs += "--logger", "trx;LogFileName=test-results.trx"
$testArgs += "--logger", "console;verbosity=minimal"

# Run tests
Write-Host "Running tests..." -ForegroundColor Green
dotnet test Tests/AsposeMcpServer.Tests.csproj $testArgs

$exitCode = $LASTEXITCODE

if ($exitCode -eq 0) {
    Write-Host ""
    Write-Host "=== All Tests Passed ===" -ForegroundColor Green
    
    # Run coverage analysis if requested
    if ($Coverage) {
        Write-Host ""
        Write-Host "=== Test Coverage Analysis ===" -ForegroundColor Cyan
        if (Test-Path "analyze-test-coverage.ps1") {
            & pwsh analyze-test-coverage.ps1
        } else {
            Write-Host "Coverage analysis script not found. Run analyze-test-coverage.ps1 separately." -ForegroundColor Yellow
        }
    }
} else {
    Write-Host ""
    Write-Host "=== Some Tests Failed ===" -ForegroundColor Red
}

# Clean up environment variable
Remove-Item Env:\SKIP_ASPOSE_LICENSE -ErrorAction SilentlyContinue

# Exit with the same exit code as dotnet test
exit $exitCode
