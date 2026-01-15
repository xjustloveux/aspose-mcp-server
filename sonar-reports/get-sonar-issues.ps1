param(
    [switch]$ListIssues,
    [switch]$ListHotspots,
    [switch]$Summary,
    [string]$FilePath,
    [ValidateSet("BUG", "VULNERABILITY", "CODE_SMELL")]
    [string]$Type,
    [ValidateSet("BLOCKER", "CRITICAL", "MAJOR", "MINOR", "INFO")]
    [string]$Severity,
    [int]$PageSize = 100
)

$ProjectKey = "xjustloveux_aspose-mcp-server"
$BaseUrl = "https://sonarcloud.io/api"

function Get-SonarIssues {
    param(
        [string]$Type,
        [string]$Severity,
        [string]$FilePath,
        [int]$PageSize
    )

    $params = @(
        "componentKeys=$ProjectKey",
        "ps=$PageSize",
        "resolved=false"
    )

    if ($Type) { $params += "types=$Type" }
    if ($Severity) { $params += "severities=$Severity" }
    if ($FilePath) { $params += "files=$FilePath" }

    $url = "$BaseUrl/issues/search?" + ($params -join "&")

    try {
        $response = Invoke-RestMethod -Uri $url -Method Get
        return $response
    }
    catch {
        Write-Error "Failed to fetch issues: $_"
        return $null
    }
}

function Get-SonarHotspots {
    param([int]$PageSize)

    $url = "$BaseUrl/hotspots/search?projectKey=$ProjectKey&ps=$PageSize"

    try {
        $response = Invoke-RestMethod -Uri $url -Method Get
        return $response
    }
    catch {
        Write-Error "Failed to fetch hotspots: $_"
        return $null
    }
}

function Get-ProjectMeasures {
    $metrics = "bugs,vulnerabilities,code_smells,coverage,duplicated_lines_density,ncloc,security_hotspots"
    $url = "$BaseUrl/measures/component?component=$ProjectKey&metricKeys=$metrics"

    try {
        $response = Invoke-RestMethod -Uri $url -Method Get
        return $response
    }
    catch {
        Write-Error "Failed to fetch measures: $_"
        return $null
    }
}

if ($Summary) {
    Write-Host "`n=== SonarCloud Project Summary ===" -ForegroundColor Cyan
    Write-Host "Project: $ProjectKey`n"

    $measures = Get-ProjectMeasures
    if ($measures) {
        Write-Host ("{0,-25} {1}" -f "Metric", "Value")
        Write-Host ("-" * 40)
        foreach ($m in $measures.component.measures) {
            $name = switch ($m.metric) {
                "bugs" { "Bugs" }
                "vulnerabilities" { "Vulnerabilities" }
                "code_smells" { "Code Smells" }
                "coverage" { "Coverage (%)" }
                "duplicated_lines_density" { "Duplicated Lines (%)" }
                "ncloc" { "Lines of Code" }
                "security_hotspots" { "Security Hotspots" }
                default { $m.metric }
            }
            Write-Host ("{0,-25} {1}" -f $name, $m.value)
        }
    }

    $issues = Get-SonarIssues -PageSize 1
    if ($issues) {
        Write-Host "`nTotal Open Issues: $($issues.total)" -ForegroundColor Yellow
    }
}
elseif ($ListIssues) {
    $issues = Get-SonarIssues -Type $Type -Severity $Severity -FilePath $FilePath -PageSize $PageSize

    if (-not $issues -or $issues.total -eq 0) {
        Write-Host "No issues found." -ForegroundColor Green
        exit 0
    }

    Write-Host "`n=== SonarCloud Issues ===" -ForegroundColor Cyan
    Write-Host "Total: $($issues.total) issues`n"

    Write-Host ("{0,-12} {1,-10} {2,-50} {3}" -f "Type", "Severity", "File", "Line")
    Write-Host ("-" * 100)

    foreach ($issue in $issues.issues) {
        $file = if ($issue.component) {
            $issue.component -replace "^${ProjectKey}:", ""
        } else { "N/A" }

        if ($file.Length -gt 48) {
            $file = "..." + $file.Substring($file.Length - 45)
        }

        $line = if ($issue.line) { $issue.line } else { "-" }

        $severityColor = switch ($issue.severity) {
            "BLOCKER" { "Red" }
            "CRITICAL" { "Red" }
            "MAJOR" { "Yellow" }
            "MINOR" { "White" }
            default { "Gray" }
        }

        Write-Host ("{0,-12} " -f $issue.type) -NoNewline
        Write-Host ("{0,-10} " -f $issue.severity) -ForegroundColor $severityColor -NoNewline
        Write-Host ("{0,-50} {1}" -f $file, $line)
    }

    if ($issues.total -gt $PageSize) {
        Write-Host "`nShowing first $PageSize of $($issues.total) issues. Use -PageSize to show more." -ForegroundColor Yellow
    }
}
elseif ($ListHotspots) {
    $hotspots = Get-SonarHotspots -PageSize $PageSize

    if (-not $hotspots -or $hotspots.paging.total -eq 0) {
        Write-Host "No security hotspots found." -ForegroundColor Green
        exit 0
    }

    Write-Host "`n=== Security Hotspots ===" -ForegroundColor Cyan
    Write-Host "Total: $($hotspots.paging.total) hotspots`n"

    Write-Host ("{0,-15} {1,-50} {2}" -f "Status", "File", "Line")
    Write-Host ("-" * 80)

    foreach ($hs in $hotspots.hotspots) {
        $file = if ($hs.component) {
            $hs.component -replace "^${ProjectKey}:", ""
        } else { "N/A" }

        if ($file.Length -gt 48) {
            $file = "..." + $file.Substring($file.Length - 45)
        }

        $line = if ($hs.line) { $hs.line } else { "-" }

        Write-Host ("{0,-15} {1,-50} {2}" -f $hs.vulnerabilityProbability, $file, $line)
    }
}
elseif ($FilePath) {
    $issues = Get-SonarIssues -FilePath $FilePath -PageSize $PageSize

    if (-not $issues -or $issues.total -eq 0) {
        Write-Host "No issues found for file: $FilePath" -ForegroundColor Green
        exit 0
    }

    Write-Host "`n=== Issues for: $FilePath ===" -ForegroundColor Cyan
    Write-Host "Total: $($issues.total) issues`n"

    foreach ($issue in $issues.issues) {
        $line = if ($issue.line) { "Line $($issue.line)" } else { "N/A" }

        $severityColor = switch ($issue.severity) {
            "BLOCKER" { "Red" }
            "CRITICAL" { "Red" }
            "MAJOR" { "Yellow" }
            "MINOR" { "White" }
            default { "Gray" }
        }

        Write-Host "[$($issue.type)] " -NoNewline -ForegroundColor Cyan
        Write-Host "$($issue.severity) " -NoNewline -ForegroundColor $severityColor
        Write-Host "@ $line"
        Write-Host "  $($issue.message)" -ForegroundColor White
        Write-Host ""
    }
}
else {
    Write-Host @"
SonarCloud Issues Query Tool

Usage:
  .\get-sonar-issues.ps1 -Summary                    # Show project summary
  .\get-sonar-issues.ps1 -ListIssues                 # List all open issues
  .\get-sonar-issues.ps1 -ListIssues -Type BUG       # List only bugs
  .\get-sonar-issues.ps1 -ListIssues -Severity MAJOR # List major issues
  .\get-sonar-issues.ps1 -ListHotspots               # List security hotspots
  .\get-sonar-issues.ps1 -FilePath "Tools/xxx.cs"    # Issues for specific file

Options:
  -Type       BUG, VULNERABILITY, CODE_SMELL
  -Severity   BLOCKER, CRITICAL, MAJOR, MINOR, INFO
  -PageSize   Number of results (default: 100)

Examples:
  # List critical bugs
  .\get-sonar-issues.ps1 -ListIssues -Type BUG -Severity CRITICAL

  # Check specific file
  .\get-sonar-issues.ps1 -FilePath "Tools/Word/DocWatermarkTool.cs"
"@
}
