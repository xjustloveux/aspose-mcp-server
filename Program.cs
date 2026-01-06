using System.Net;
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

try
{
    config.Validate();
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
    var host = transportConfig.Mode switch
    {
        TransportMode.Stdio => CreateStdioHost(),
        TransportMode.Sse => CreateSseHost(transportConfig),
        TransportMode.WebSocket => CreateWebSocketHost(transportConfig),
        _ => throw new ArgumentException($"Unknown transport mode: {transportConfig.Mode}")
    };

    await host.RunAsync();
}
catch (Exception ex)
{
    Console.Error.WriteLine($"[ERROR] Fatal error: {ex.GetType().Name}");
#if DEBUG
    Console.Error.WriteLine($"[ERROR] Details: {ex.Message}");
    Console.Error.WriteLine($"[ERROR] Stack trace: {ex.StackTrace}");
#else
    Console.Error.WriteLine($"[ERROR] An internal error occurred. Check logs for details.");
#endif
    Environment.Exit(1);
}

// Creates a host configured for stdio transport mode
IHost CreateStdioHost()
{
    var builder = Host.CreateApplicationBuilder(args);

    builder.Logging.ClearProviders();
    builder.Logging.AddConsole(options => { options.LogToStandardErrorThreshold = LogLevel.Trace; });

    // Register configurations
    builder.Services.AddSingleton(transportConfig);
    builder.Services.AddSingleton(sessionConfig);
    builder.Services.AddSingleton(authConfig);
    builder.Services.AddSingleton(authConfig.ApiKey);
    builder.Services.AddSingleton(authConfig.Jwt);
    builder.Services.AddSingleton(trackingConfig);
    builder.Services.AddSingleton<DocumentSessionManager>();
    builder.Services.AddSingleton<TempFileManager>();
    builder.Services.AddHostedService(sp => sp.GetRequiredService<TempFileManager>());

    // Register session identity accessor for Stdio mode
    // (reads from environment variables if available, otherwise anonymous)
    builder.Services.AddSingleton<ISessionIdentityAccessor, StdioSessionIdentityAccessor>();

    builder.Services.AddMcpServer()
        .WithStdioServerTransport()
        .WithFilteredTools(config, sessionConfig);

    return builder.Build();
}

// Creates a host configured for Server-Sent Events (SSE) transport mode
IHost CreateSseHost(TransportConfig transport)
{
    var builder = WebApplication.CreateBuilder(args);

    builder.WebHost.ConfigureKestrel(options =>
    {
        if (transport.Host == "localhost")
            options.ListenLocalhost(transport.Port);
        else if (transport.Host == "0.0.0.0" || transport.Host == "*")
            options.ListenAnyIP(transport.Port);
        else
            options.Listen(IPAddress.Parse(transport.Host), transport.Port);
    });

    builder.Logging.ClearProviders();
    builder.Logging.AddConsole(options => { options.LogToStandardErrorThreshold = LogLevel.Trace; });

    // Register configurations
    builder.Services.AddSingleton(transportConfig);
    builder.Services.AddSingleton(sessionConfig);
    builder.Services.AddSingleton(authConfig);
    builder.Services.AddSingleton(authConfig.ApiKey);
    builder.Services.AddSingleton(authConfig.Jwt);
    builder.Services.AddSingleton(trackingConfig);
    builder.Services.AddSingleton<DocumentSessionManager>();
    builder.Services.AddSingleton<TempFileManager>();
    builder.Services.AddHostedService(sp => sp.GetRequiredService<TempFileManager>());

    // Register session identity accessor for SSE mode (from HttpContext)
    builder.Services.AddHttpContextAccessor();
    builder.Services.AddSingleton<ISessionIdentityAccessor, HttpContextSessionIdentityAccessor>();

    builder.Services.AddMcpServer()
        .WithFilteredTools(config, sessionConfig);

    var app = builder.Build();

    Console.Error.WriteLine($"[INFO] SSE server listening on http://{transport.Host}:{transport.Port}");

    // Add authentication middleware if enabled
    if (authConfig.ApiKey.Enabled)
    {
        Console.Error.WriteLine($"[INFO] API Key authentication enabled (mode: {authConfig.ApiKey.Mode})");
        app.UseMiddleware<ApiKeyAuthenticationMiddleware>();
    }

    if (authConfig.Jwt.Enabled)
    {
        Console.Error.WriteLine($"[INFO] JWT authentication enabled (mode: {authConfig.Jwt.Mode})");
        app.UseMiddleware<JwtAuthenticationMiddleware>();
    }

    // Add tracking middleware if enabled
    if (trackingConfig.LogEnabled || trackingConfig.WebhookEnabled || trackingConfig.MetricsEnabled)
    {
        Console.Error.WriteLine("[INFO] Tracking middleware enabled");
        app.UseMiddleware<TrackingMiddleware>();
    }

    // Health check endpoints for container orchestration
    app.MapGet("/health", () => Results.Ok(new { status = "healthy" }));
    app.MapGet("/ready", () => Results.Ok(new { status = "ready" }));

    app.MapMcp("/mcp");

    return app;
}

// Creates a host configured for WebSocket transport mode
IHost CreateWebSocketHost(TransportConfig transport)
{
    var builder = WebApplication.CreateBuilder(args);

    builder.WebHost.ConfigureKestrel(options =>
    {
        if (transport.Host == "localhost")
            options.ListenLocalhost(transport.Port);
        else if (transport.Host == "0.0.0.0" || transport.Host == "*")
            options.ListenAnyIP(transport.Port);
        else
            options.Listen(IPAddress.Parse(transport.Host), transport.Port);
    });

    builder.Logging.ClearProviders();
    builder.Logging.AddConsole(options => { options.LogToStandardErrorThreshold = LogLevel.Trace; });

    // Register configurations
    builder.Services.AddSingleton(transportConfig);
    builder.Services.AddSingleton(sessionConfig);
    builder.Services.AddSingleton(authConfig);
    builder.Services.AddSingleton(authConfig.ApiKey);
    builder.Services.AddSingleton(authConfig.Jwt);
    builder.Services.AddSingleton(trackingConfig);
    builder.Services.AddSingleton<DocumentSessionManager>();
    builder.Services.AddSingleton<TempFileManager>();
    builder.Services.AddHostedService(sp => sp.GetRequiredService<TempFileManager>());

    // Register session identity accessor for WebSocket mode (from HttpContext)
    builder.Services.AddHttpContextAccessor();
    builder.Services.AddSingleton<ISessionIdentityAccessor, HttpContextSessionIdentityAccessor>();

    var app = builder.Build();

    Console.Error.WriteLine($"[INFO] WebSocket server listening on ws://{transport.Host}:{transport.Port}");

    // Add authentication middleware if enabled
    if (authConfig.ApiKey.Enabled)
    {
        Console.Error.WriteLine($"[INFO] API Key authentication enabled (mode: {authConfig.ApiKey.Mode})");
        app.UseMiddleware<ApiKeyAuthenticationMiddleware>();
    }

    if (authConfig.Jwt.Enabled)
    {
        Console.Error.WriteLine($"[INFO] JWT authentication enabled (mode: {authConfig.Jwt.Mode})");
        app.UseMiddleware<JwtAuthenticationMiddleware>();
    }

    // Add tracking middleware if enabled
    if (trackingConfig.LogEnabled || trackingConfig.WebhookEnabled || trackingConfig.MetricsEnabled)
    {
        Console.Error.WriteLine("[INFO] Tracking middleware enabled");
        app.UseMiddleware<TrackingMiddleware>();
    }

    app.UseWebSockets();

    // Health check endpoints for container orchestration
    app.MapGet("/health", () => Results.Ok(new { status = "healthy" }));
    app.MapGet("/ready", () => Results.Ok(new { status = "ready" }));

    var executablePath = Environment.ProcessPath ?? "dotnet";
    var toolArgs = string.Join(" ", args.Where(a =>
        a.StartsWith("--word", StringComparison.OrdinalIgnoreCase) ||
        a.StartsWith("--excel", StringComparison.OrdinalIgnoreCase) ||
        a.StartsWith("--powerpoint", StringComparison.OrdinalIgnoreCase) ||
        a.StartsWith("--ppt", StringComparison.OrdinalIgnoreCase) ||
        a.StartsWith("--pdf", StringComparison.OrdinalIgnoreCase) ||
        a.StartsWith("--all", StringComparison.OrdinalIgnoreCase) ||
        a.StartsWith("--license", StringComparison.OrdinalIgnoreCase) ||
        a.StartsWith("--session", StringComparison.OrdinalIgnoreCase)));

    var handler = new WebSocketConnectionHandler(executablePath, toolArgs, app.Services.GetService<ILoggerFactory>());

    app.Map("/ws", async context =>
    {
        if (context.WebSockets.IsWebSocketRequest)
        {
            // Extract authentication context set by auth middleware
            var tenantId = context.Items["TenantId"]?.ToString();
            var userId = context.Items["UserId"]?.ToString();

            var webSocket = await context.WebSockets.AcceptWebSocketAsync();
            await handler.HandleConnectionAsync(webSocket, context.RequestAborted, tenantId, userId);
        }
        else
        {
            context.Response.StatusCode = 400;
        }
    });

    return app;
}