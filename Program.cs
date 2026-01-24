using System.Text;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Security;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Core.Tracking;
using AsposeMcpServer.Core.Transport;

Console.OutputEncoding = Encoding.UTF8;
Console.InputEncoding = Encoding.UTF8;

var config = ServerConfig.LoadFromArgs(args);
var transportConfig = TransportConfig.LoadFromArgs(args);
var sessionConfig = SessionConfig.LoadFromArgs(args);
var authConfig = AuthConfig.LoadFromArgs(args);
var trackingConfig = TrackingConfig.LoadFromArgs(args);
var originConfig = OriginValidationConfig.LoadFromArgs(args);

try
{
    config.Validate();
    transportConfig.Validate();
    sessionConfig.Validate();
    authConfig.Validate();
    trackingConfig.Validate();
}
catch (InvalidOperationException ex)
{
    Console.Error.WriteLine($"[ERROR] Configuration error: {ex.Message}");
    Environment.Exit(1);
}

var toolFilter = new ToolFilterService(config, sessionConfig);
Console.Error.WriteLine($"[INFO] Aspose MCP Server - Enabled categories: {toolFilter.GetEnabledCategories()}");
Console.Error.WriteLine($"[INFO] Transport mode: {transportConfig.Mode}");
if (sessionConfig.Enabled)
    Console.Error.WriteLine($"[INFO] Session isolation mode: {sessionConfig.IsolationMode}");
await Console.Error.FlushAsync();

LicenseManager.SetLicense(config);

try
{
    var host = HostFactory.CreateHost(args, config, transportConfig, sessionConfig, authConfig, trackingConfig,
        originConfig);
    await host.RunAsync();
}
catch (Exception ex)
{
    Console.Error.WriteLine($"[ERROR] Fatal error: {ex.GetType().Name}");
#if DEBUG
    Console.Error.WriteLine($"[ERROR] Details: {ex.Message}");
    Console.Error.WriteLine($"[ERROR] Stack trace: {ex.StackTrace}");
#else
    Console.Error.WriteLine("[ERROR] An internal error occurred. Check logs for details.");
#endif
    Environment.Exit(1);
}
