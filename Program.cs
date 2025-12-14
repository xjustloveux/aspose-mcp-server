using System.Text;
using AsposeMcpServer;
using AsposeMcpServer.Core;

try
{
    // Set console encoding to UTF-8 for proper Chinese character support
    Console.OutputEncoding = Encoding.UTF8;
    Console.InputEncoding = Encoding.UTF8;
    
    // Load configuration from command line arguments
    var config = ServerConfig.LoadFromArgs(args);
    
    // Validate configuration
    try
    {
        config.Validate();
    }
    catch (InvalidOperationException ex)
    {
        Console.Error.WriteLine($"[ERROR] Configuration error: {ex.Message}");
        Environment.Exit(1);
    }
    
    Console.Error.WriteLine($"[INFO] Aspose MCP Server - Enabled tools: {config.GetEnabledToolsInfo()}");
    await Console.Error.FlushAsync();
    
    // Initialize Aspose License (with stdout suppression)
    LicenseManager.SetLicense(config);

    // Create and run MCP server with configuration
    var server = new McpServer(config);
    await server.RunAsync();
}
catch (Exception ex)
{
    // All errors to stderr, never stdout
    // Don't expose full stack trace in production to prevent information leakage
    Console.Error.WriteLine($"[ERROR] Fatal error: {ex.GetType().Name}");
#if DEBUG
    Console.Error.WriteLine($"[ERROR] Details: {ex.Message}");
    Console.Error.WriteLine($"[ERROR] Stack trace: {ex.StackTrace}");
#else
    Console.Error.WriteLine($"[ERROR] An internal error occurred. Check logs for details.");
#endif
    Environment.Exit(1);
}

