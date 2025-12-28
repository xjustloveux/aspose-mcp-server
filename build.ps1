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
dotnet build AsposeMcpServer.csproj --configuration $configuration

# Exit with the same exit code as dotnet build
exit $LASTEXITCODE

