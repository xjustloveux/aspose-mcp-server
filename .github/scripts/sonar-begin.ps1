<#
.SYNOPSIS
    Begins a SonarCloud analysis session with all project-specific settings.

.PARAMETER Token
    SonarCloud authentication token.

.PARAMETER ScannerPath
    Path to the dotnet-sonarscanner executable.
    Use "dotnet sonarscanner" for global tool, or ".\.sonar\scanner\dotnet-sonarscanner" for local tool.
#>
param(
    [Parameter(Mandatory)]
    [string]$Token,

    [Parameter(Mandatory)]
    [string]$ScannerPath
)

$ErrorActionPreference = "Stop"

$scannerArgs = @(
    "begin",
    "/k:xjustloveux_aspose-mcp-server",
    "/o:xjustloveux",
    "/d:sonar.host.url=https://sonarcloud.io",
    "/d:sonar.token=$Token",
    "/d:sonar.sourceEncoding=UTF-8",
    "/d:sonar.cs.opencover.reportsPaths=**/coverage.opencover.xml",
    "/d:sonar.coverage.exclusions=**/Tests/**,**/Program.cs",
    "/d:sonar.exclusions=**/deploy/web.config,**/deploy/deployment.yaml",
    "/d:sonar.cpd.exclusions=**/Results/**",
    "/d:sonar.issue.ignore.multicriteria=e1,e2,e3,e4,e5,e6,e7,e8,e9,e10,e11,e12,e13,e14,e15,e16,e17,e18,e19,e20,e21,e22,e23,e24,e25,e26,e27,e28,e29,e30,e31,e32,e33,e34,e35,e36",
    "/d:sonar.issue.ignore.multicriteria.e1.ruleKey=csharpsquid:S3973",
    "/d:sonar.issue.ignore.multicriteria.e1.resourceKey=**/*",
    "/d:sonar.issue.ignore.multicriteria.e2.ruleKey=docker:S7020",
    "/d:sonar.issue.ignore.multicriteria.e2.resourceKey=**/Dockerfile",
    "/d:sonar.issue.ignore.multicriteria.e3.ruleKey=csharpsquid:S107",
    "/d:sonar.issue.ignore.multicriteria.e3.resourceKey=**/Tools/**",
    "/d:sonar.issue.ignore.multicriteria.e4.ruleKey=csharpsquid:S1144",
    "/d:sonar.issue.ignore.multicriteria.e4.resourceKey=**/ApiKeyAuthenticationMiddleware.cs",
    "/d:sonar.issue.ignore.multicriteria.e5.ruleKey=csharpsquid:S1144",
    "/d:sonar.issue.ignore.multicriteria.e5.resourceKey=**/JwtAuthenticationMiddleware.cs",
    "/d:sonar.issue.ignore.multicriteria.e6.ruleKey=csharpsquid:S3459",
    "/d:sonar.issue.ignore.multicriteria.e6.resourceKey=**/ApiKeyAuthenticationMiddleware.cs",
    "/d:sonar.issue.ignore.multicriteria.e7.ruleKey=csharpsquid:S3459",
    "/d:sonar.issue.ignore.multicriteria.e7.resourceKey=**/JwtAuthenticationMiddleware.cs",
    "/d:sonar.issue.ignore.multicriteria.e8.ruleKey=docker:S6570",
    "/d:sonar.issue.ignore.multicriteria.e8.resourceKey=**/Dockerfile",
    "/d:sonar.issue.ignore.multicriteria.e9.ruleKey=docker:S7031",
    "/d:sonar.issue.ignore.multicriteria.e9.resourceKey=**/Dockerfile",
    "/d:sonar.issue.ignore.multicriteria.e10.ruleKey=csharpsquid:S3776",
    "/d:sonar.issue.ignore.multicriteria.e10.resourceKey=**/*",
    "/d:sonar.issue.ignore.multicriteria.e11.ruleKey=csharpsquid:S2368",
    "/d:sonar.issue.ignore.multicriteria.e11.resourceKey=**/Tools/**",
    "/d:sonar.issue.ignore.multicriteria.e12.ruleKey=csharpsquid:S3011",
    "/d:sonar.issue.ignore.multicriteria.e12.resourceKey=**/McpServerBuilderExtensions.cs",
    "/d:sonar.issue.ignore.multicriteria.e13.ruleKey=csharpsquid:S3881",
    "/d:sonar.issue.ignore.multicriteria.e13.resourceKey=**/DocumentContext.cs",
    "/d:sonar.issue.ignore.multicriteria.e14.ruleKey=csharpsquid:S2925",
    "/d:sonar.issue.ignore.multicriteria.e14.resourceKey=**/Tests/**",
    "/d:sonar.issue.ignore.multicriteria.e15.ruleKey=csharpsquid:S3267",
    "/d:sonar.issue.ignore.multicriteria.e15.resourceKey=**/AsposeHelper.cs",
    "/d:sonar.issue.ignore.multicriteria.e16.ruleKey=csharpsquid:S1192",
    "/d:sonar.issue.ignore.multicriteria.e16.resourceKey=**/ConvertDocumentTool.cs",
    "/d:sonar.issue.ignore.multicriteria.e17.ruleKey=docker:S6470",
    "/d:sonar.issue.ignore.multicriteria.e17.resourceKey=**/Dockerfile",
    "/d:sonar.issue.ignore.multicriteria.e18.ruleKey=githubactions:S7637",
    "/d:sonar.issue.ignore.multicriteria.e18.resourceKey=**/.github/workflows/*.yml",
    "/d:sonar.issue.ignore.multicriteria.e19.ruleKey=githubactions:S7636",
    "/d:sonar.issue.ignore.multicriteria.e19.resourceKey=**/test.yml",
    "/d:sonar.issue.ignore.multicriteria.e20.ruleKey=csharpsquid:S4790",
    "/d:sonar.issue.ignore.multicriteria.e20.resourceKey=**/PptImageHelper.cs",
    "/d:sonar.issue.ignore.multicriteria.e21.ruleKey=csharpsquid:S5443",
    "/d:sonar.issue.ignore.multicriteria.e21.resourceKey=**/SessionConfig.cs",
    "/d:sonar.issue.ignore.multicriteria.e22.ruleKey=csharpsquid:S6667",
    "/d:sonar.issue.ignore.multicriteria.e22.resourceKey=**/WebSocketConnectionHandler.cs",
    "/d:sonar.issue.ignore.multicriteria.e23.ruleKey=csharpsquid:S6667",
    "/d:sonar.issue.ignore.multicriteria.e23.resourceKey=**/TrackingMiddleware.cs",
    "/d:sonar.issue.ignore.multicriteria.e24.ruleKey=csharpsquid:S6667",
    "/d:sonar.issue.ignore.multicriteria.e24.resourceKey=**/JwtAuthenticationMiddleware.cs",
    "/d:sonar.issue.ignore.multicriteria.e25.ruleKey=csharpsquid:S6966",
    "/d:sonar.issue.ignore.multicriteria.e25.resourceKey=**/WebSocketConnectionHandler.cs",
    "/d:sonar.issue.ignore.multicriteria.e26.ruleKey=csharpsquid:S107",
    "/d:sonar.issue.ignore.multicriteria.e26.resourceKey=**/FontHelper.cs",
    "/d:sonar.issue.ignore.multicriteria.e27.ruleKey=csharpsquid:S127",
    "/d:sonar.issue.ignore.multicriteria.e27.resourceKey=**/TransportConfig.cs",
    "/d:sonar.issue.ignore.multicriteria.e28.ruleKey=csharpsquid:S1854",
    "/d:sonar.issue.ignore.multicriteria.e28.resourceKey=**/TrackingMiddleware.cs",
    "/d:sonar.issue.ignore.multicriteria.e29.ruleKey=csharpsquid:S1075",
    "/d:sonar.issue.ignore.multicriteria.e29.resourceKey=**/TrackingConfig.cs",
    "/d:sonar.issue.ignore.multicriteria.e30.ruleKey=csharpsquid:S1172",
    "/d:sonar.issue.ignore.multicriteria.e30.resourceKey=**/ApiKeyAuthenticationMiddleware.cs",
    "/d:sonar.issue.ignore.multicriteria.e31.ruleKey=csharpsquid:S3881",
    "/d:sonar.issue.ignore.multicriteria.e31.resourceKey=**/TempFileManager.cs",
    "/d:sonar.issue.ignore.multicriteria.e32.ruleKey=csharpsquid:S3881",
    "/d:sonar.issue.ignore.multicriteria.e32.resourceKey=**/DocumentSessionManager.cs",
    "/d:sonar.issue.ignore.multicriteria.e33.ruleKey=csharpsquid:S3267",
    "/d:sonar.issue.ignore.multicriteria.e33.resourceKey=**/SetColumnWidthWordTableHandler.cs",
    "/d:sonar.issue.ignore.multicriteria.e34.ruleKey=csharpsquid:S3267",
    "/d:sonar.issue.ignore.multicriteria.e34.resourceKey=**/DeleteColumnWordTableHandler.cs",
    "/d:sonar.issue.ignore.multicriteria.e35.ruleKey=csharpsquid:S3267",
    "/d:sonar.issue.ignore.multicriteria.e35.resourceKey=**/GetWordListFormatHandler.cs",
    "/d:sonar.issue.ignore.multicriteria.e36.ruleKey=csharpsquid:S2094",
    "/d:sonar.issue.ignore.multicriteria.e36.resourceKey=**/RunFormatInfo.cs"
)

& $ScannerPath @scannerArgs

if ($LASTEXITCODE -ne 0) {
    throw "SonarScanner begin failed with exit code $LASTEXITCODE"
}
