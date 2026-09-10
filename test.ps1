# PowerShell script to run unit tests with UTF-8 encoding
# This script sets UTF-8 encoding to prevent Chinese character display issues

param(
    [switch]$Verbose,
    [switch]$NoBuild,
    [switch]$Coverage,
    [string]$Filter,
    [switch]$SkipLicense,  # Skip license loading, force evaluation mode
    # A run that executes nothing is not a pass. This says one was expected — a workflow that
    # deliberately selects only licensed-only tests in evaluation mode, say — and it reports
    # "NO TESTS EXECUTED" and exits 0 rather than claiming everything passed (R17-T01).
    [switch]$AllowNoTests,
    [string]$Configuration = "Release",  # Build configuration (Debug or Release)
    [string]$LogFile,  # Output failed test details to file (e.g. -LogFile "Tests\TestResults\log.txt")
    # Name of the TRX this run writes. Every run used to write test-results.trx, so the evidence
    # for one run was gone as soon as the next one started and a failure could not be attributed
    # to the run that produced it (R8-DOC01). Defaults to the old name so existing callers are
    # unaffected; pass a unique one when the result has to survive.
    [string]$TrxName = "test-results.trx"
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
# Always specify configuration to ensure correct DLL location
$testArgs += "--configuration", $Configuration

if ($Coverage) {
    $testArgs += "--collect", "XPlat Code Coverage"
}

if ($Filter) {
    $testArgs += "--filter", $Filter
}

# Add logger for test results
$testArgs += "--logger", "trx;LogFileName=$TrxName"
$testArgs += "--logger", "console;verbosity=minimal"

# The TRX name is reusable, so a run that writes no logger output would leave the previous
# run's file in place — and every count read back from it would describe that run instead
# (R10-T02). The identity of whatever is there now is recorded before the run starts.
$trxPath = Join-Path "Tests/TestResults" $TrxName
$trxBefore = if (Test-Path $trxPath) { (Get-Item $trxPath).LastWriteTimeUtc } else { $null }
$runStartedUtc = (Get-Date).ToUniversalTime()

# Run tests
Write-Host "Running tests..." -ForegroundColor Green
dotnet test Tests/AsposeMcpServer.Tests.csproj $testArgs

$exitCode = $LASTEXITCODE

# A --filter that matches no test still exits 0, so a targeted run that executed nothing
# printed "All Tests Passed" and every gate reported from that line was reporting on an empty
# run. The count is read back from the TRX this run just wrote, which is the only record of
# what actually executed.
# Whether the TRX at $trxPath was written by this run. Both branches ask: the success branch
# reads counts from it, and the failure branch exports failed tests from it. Asking only on
# success meant a runner that failed before writing anything exported the *previous* run's
# failures as this run's report (§23.6).
$trxIsFromThisRun = $false
if (Test-Path $trxPath) {
    $trxAfter = (Get-Item $trxPath).LastWriteTimeUtc
    $trxIsFromThisRun = -not (($null -ne $trxBefore -and $trxAfter -le $trxBefore) -or ($trxAfter -lt $runStartedUtc))
}

if ($exitCode -eq 0) {
    if (-not (Test-Path $trxPath)) {
        Write-Host ""
        Write-Host "=== No Test Results ===" -ForegroundColor Red
        Write-Host "dotnet test reported success but wrote no $TrxName, so nothing is known about what ran." -ForegroundColor Red
        exit 1
    }

    if (-not $trxIsFromThisRun) {
        Write-Host ""
        Write-Host "=== Stale Test Results ===" -ForegroundColor Red
        Write-Host "$TrxName was not written by this run, so its counts describe a different execution." -ForegroundColor Red
        exit 1
    }

    $counters = ([xml](Get-Content $trxPath)).TestRun.ResultSummary.Counters
    $total = [int]$counters.total
    $executed = [int]$counters.executed

    # `executed`, not `total`. `total` counts skipped tests, so a run where every match declined to
    # run reported 12 total, exited 0, and printed the green line — the same false green as a
    # filter matching nothing, one step along (R17-T01). A run that executed nothing proves
    # nothing, whether the tests were absent or present and skipped.
    if ($total -eq 0 -or $executed -eq 0) {
        Write-Host ""
        Write-Host "=== NO TESTS EXECUTED ===" -ForegroundColor Red
        if ($total -eq 0) {
            if ($Filter) {
                Write-Host "The filter '$Filter' matched no test. An empty run is not a pass." `
                    -ForegroundColor Red
            } else {
                Write-Host "The run executed no tests at all. An empty run is not a pass." `
                    -ForegroundColor Red
            }
        } else {
            $subject = if ($Filter) { "The filter '$Filter' matched" } else { "This run matched" }
            Write-Host "$subject $total test(s) and every one of them was skipped, so nothing was executed. An all-skipped run is not a pass." `
                -ForegroundColor Red
            Write-Host "If a run of only skipped tests is what you meant, say so with -AllowNoTests." `
                -ForegroundColor Yellow
        }

        if (-not $AllowNoTests) { exit 1 }

        Write-Host "-AllowNoTests was given, so this is reported as a skipped run rather than a pass." `
            -ForegroundColor Yellow
        exit 0
    }
}

if ($exitCode -eq 0) {
    Write-Host ""
    Write-Host "=== All Tests Passed ===" -ForegroundColor Green
    
    # Run coverage analysis if requested
    if ($Coverage) {
        Write-Host ""
        Write-Host "=== Test Coverage Analysis ===" -ForegroundColor Cyan
        # Coverage data is collected above by "--collect XPlat Code Coverage"; the report itself
        # is produced by Codecov in CI. There is no analyze-test-coverage.ps1 in this repository,
        # so pointing at one told the reader to run a file that does not exist.
        $coverageFiles = Get-ChildItem -Path "Tests/TestResults" -Filter "coverage.cobertura.xml" -Recurse -ErrorAction SilentlyContinue
        if ($coverageFiles) {
            Write-Host "Coverage collected:" -ForegroundColor Green
            foreach ($file in $coverageFiles) { Write-Host "  $($file.FullName)" -ForegroundColor Gray }
            Write-Host "See coverage-reports/get-uncovered-lines.ps1 to query uncovered lines." -ForegroundColor Gray
        } else {
            Write-Host "No coverage file was produced." -ForegroundColor Yellow
        }
    }
} else {
    Write-Host ""
    Write-Host "=== Some Tests Failed ===" -ForegroundColor Red

    # Export failed test details to file if LogFile parameter is specified
    if ($LogFile) {
        # This run wrote $trxPath. Picking the newest TRX in the directory could report a
        # different run's failures entirely, which is how evidence gets attributed to the
        # wrong execution (§21.8).
        # Only this run's artifact. Exporting a stale one reported test cases that never ran in
        # this execution, under this run's heading (§23.6).
        $trxFile = if ($trxIsFromThisRun) { Get-Item $trxPath -ErrorAction SilentlyContinue } else { $null }
        if (-not $trxFile) {
            Write-Host "No test results from this run to export; reporting the runner's exit code only." -ForegroundColor Yellow
        }
        if ($trxFile) {
            $xml = [xml](Get-Content $trxFile.FullName)
            $ns = @{ t = "http://microsoft.com/schemas/VisualStudio/TeamTest/2010" }
            $failed = $xml | Select-Xml "//t:UnitTestResult[@outcome='Failed']" -Namespace $ns

            $output = @()
            $output += "=== Failed Tests Report ==="
            $output += "Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
            $output += "Total Failed: $($failed.Count)"
            $output += ""

            foreach ($result in $failed) {
                $node = $result.Node
                $testName = $node.testName
                $errorMsg = $node.Output.ErrorInfo.Message
                $stackTrace = $node.Output.ErrorInfo.StackTrace
                $output += "--- $testName ---"
                $output += "Error: $errorMsg"
                if ($stackTrace) {
                    $output += "Stack: $($stackTrace.Split("`n") | Select-Object -First 3 | ForEach-Object { $_.Trim() })"
                }
                $output += ""
            }

            $logDir = Split-Path $LogFile -Parent
            if ($logDir -and !(Test-Path $logDir)) {
                New-Item -ItemType Directory -Path $logDir -Force | Out-Null
            }
            $output | Out-File -FilePath $LogFile -Encoding utf8
            Write-Host "Failed test details saved to: $LogFile" -ForegroundColor Yellow
        } else {
            Write-Host "No .trx file found to extract failures from." -ForegroundColor Yellow
        }
    }
}

# Clean up environment variable
Remove-Item Env:\SKIP_ASPOSE_LICENSE -ErrorAction SilentlyContinue

# Exit with the same exit code as dotnet test
exit $exitCode
