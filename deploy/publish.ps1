# Aspose MCP Server - Cross-platform Build Script

param(
    [switch]$Windows,
    [switch]$Linux,
    [switch]$MacOS,
    [switch]$IIS,
    [switch]$All,
    [switch]$Clean
)

$ErrorActionPreference = "Stop"

Write-Host "=== Aspose MCP Server - Cross-platform Build ===" -ForegroundColor Cyan
Write-Host ""

# Clean output directory
if ($Clean -or $All) {
    Write-Host "Cleaning output directory..." -ForegroundColor Yellow
    if (Test-Path "publish") {
        Remove-Item -Path "publish" -Recurse -Force
    }
}

# Create output directory
New-Item -ItemType Directory -Force -Path "publish" | Out-Null

function Build-Platform {
    param(
        [string]$Runtime,
        [string]$Platform
    )
    
    Write-Host "Building for $Platform ($Runtime)..." -ForegroundColor Green
    
    $outputPath = "publish/$Platform"
    
    # Get version from Git tag if available, otherwise use default
    $version = if ($env:VERSION) { 
        $env:VERSION 
    } else {
        try {
            $gitTag = git describe --tags --abbrev=0 2>&1 | Select-Object -First 1
            if ($gitTag -and -not ($gitTag -is [System.Management.Automation.ErrorRecord])) {
                $gitTag.ToString().TrimStart('v')
            } else {
                "1.0.0"
            }
        } catch {
            "1.0.0"
        }
    }
    
    Write-Host "  Using version: $version" -ForegroundColor Gray
    
    dotnet publish AsposeMcpServer.csproj `
        --configuration Release `
        --runtime $Runtime `
        --self-contained true `
        --output $outputPath `
        -p:Version=$version `
        -p:PublishSingleFile=true `
        -p:PublishTrimmed=false `
        -p:IncludeNativeLibrariesForSelfExtract=true `
        -p:DebugType=none `
        --nologo `
        --verbosity quiet

    if ($LASTEXITCODE -eq 0) {
        Write-Host "  ✓ Build successful: $outputPath" -ForegroundColor Green

        # Clean up unnecessary files for standalone deployment
        $unnecessaryFiles = @(
            "*.pdb",
            "*.staticwebassets.endpoints.json",
            "web.config",
            "config_example.json"
        )
        foreach ($pattern in $unnecessaryFiles) {
            Get-ChildItem -Path $outputPath -Filter $pattern -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
        }
        # Remove deploy folder if accidentally included
        $deployFolder = Join-Path $outputPath "deploy"
        if (Test-Path $deployFolder) {
            Remove-Item -Path $deployFolder -Recurse -Force -ErrorAction SilentlyContinue
        }
        Write-Host "  ✓ Cleaned up unnecessary files" -ForegroundColor Green

        # Get directory size
        $size = (Get-ChildItem -Path $outputPath -Recurse | Measure-Object -Property Length -Sum).Sum / 1MB
        Write-Host "  Size: $([math]::Round($size, 2)) MB" -ForegroundColor Gray
    } else {
        Write-Host "  ✗ Build failed" -ForegroundColor Red
    }
    Write-Host ""
}

# Build for selected platforms
if ($All -or $Windows) {
    Build-Platform "win-x64" "windows-x64"
}

if ($All -or $Linux) {
    Build-Platform "linux-x64" "linux-x64"
}

if ($All -or $MacOS) {
    Build-Platform "osx-x64" "macos-x64"
    Build-Platform "osx-arm64" "macos-arm64"
}

if ($IIS) {
    Write-Host "Building for IIS deployment..." -ForegroundColor Green

    $outputPath = "publish/iis"

    # Get version from Git tag if available, otherwise use default
    $version = if ($env:VERSION) {
        $env:VERSION
    } else {
        try {
            $gitTag = git describe --tags --abbrev=0 2>&1 | Select-Object -First 1
            if ($gitTag -and -not ($gitTag -is [System.Management.Automation.ErrorRecord])) {
                $gitTag.ToString().TrimStart('v')
            } else {
                "1.0.0"
            }
        } catch {
            "1.0.0"
        }
    }

    Write-Host "  Using version: $version" -ForegroundColor Gray

    # IIS deployment: Not single file, not self-contained (uses shared framework)
    dotnet publish AsposeMcpServer.csproj `
        --configuration Release `
        --runtime win-x64 `
        --self-contained false `
        --output $outputPath `
        -p:Version=$version `
        -p:PublishSingleFile=false `
        --nologo `
        --verbosity quiet

    if ($LASTEXITCODE -eq 0) {
        Write-Host "  ✓ Build successful: $outputPath" -ForegroundColor Green

        # Copy web.config (relative to this script in deploy/)
        $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
        $webConfigSource = Join-Path $scriptDir "web.config"
        if (Test-Path $webConfigSource) {
            Copy-Item $webConfigSource -Destination $outputPath -Force
            Write-Host "  ✓ web.config copied" -ForegroundColor Green
        } else {
            Write-Host "  ! web.config not found at $webConfigSource" -ForegroundColor Yellow
        }

        # Get directory size
        $size = (Get-ChildItem -Path $outputPath -Recurse | Measure-Object -Property Length -Sum).Sum / 1MB
        Write-Host "  Size: $([math]::Round($size, 2)) MB" -ForegroundColor Gray
    } else {
        Write-Host "  ✗ Build failed" -ForegroundColor Red
    }
    Write-Host ""
}

# If no platform specified, show help
if (-not ($Windows -or $Linux -or $MacOS -or $IIS -or $All)) {
    Write-Host "Usage:" -ForegroundColor Yellow
    Write-Host "  .\publish.ps1 -Windows    # Build for Windows (single exe)" -ForegroundColor Gray
    Write-Host "  .\publish.ps1 -Linux      # Build for Linux" -ForegroundColor Gray
    Write-Host "  .\publish.ps1 -MacOS      # Build for macOS (Intel + ARM)" -ForegroundColor Gray
    Write-Host "  .\publish.ps1 -IIS        # Build for IIS deployment" -ForegroundColor Gray
    Write-Host "  .\publish.ps1 -All        # Build for all platforms" -ForegroundColor Gray
    Write-Host "  .\publish.ps1 -Clean      # Clean before build" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Example:" -ForegroundColor Yellow
    Write-Host "  .\publish.ps1 -All -Clean" -ForegroundColor Gray
    exit 0
}

Write-Host "=== Build Complete ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "Output directory: $(Resolve-Path 'publish')" -ForegroundColor Green
Write-Host ""
Write-Host "Usage examples:" -ForegroundColor Yellow
Write-Host ""
Write-Host "Windows:" -ForegroundColor Cyan
Write-Host '  "C:\path\to\publish\windows-x64\AsposeMcpServer.exe" --word' -ForegroundColor Gray
Write-Host ""
Write-Host "Linux/macOS:" -ForegroundColor Cyan
Write-Host '  /path/to/publish/linux-x64/AsposeMcpServer --word' -ForegroundColor Gray
Write-Host ""
Write-Host "Claude Desktop config.json:" -ForegroundColor Cyan
Write-Host '{
  "mcpServers": {
    "aspose-word": {
      "command": "C:/path/to/AsposeMcpServer.exe",
      "args": ["--word"]
    },
    "aspose-excel": {
      "command": "C:/path/to/AsposeMcpServer.exe",
      "args": ["--excel"]
    }
  }
}' -ForegroundColor Gray

