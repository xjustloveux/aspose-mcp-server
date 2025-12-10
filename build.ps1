# PowerShell script to build with UTF-8 encoding
# This script sets UTF-8 encoding to prevent Chinese character display issues

# Set console encoding to UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# Set environment variable for .NET
$env:DOTNET_CLI_UI_LANGUAGE = "en-US"

# Run dotnet build
dotnet build $args

# Exit with the same exit code as dotnet build
exit $LASTEXITCODE

