// NOSONAR S6966 - Top-level statements entry point; sync Console.Error for startup logging

using System.Net;
using System.Text;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Security;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Core.Tasks;
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
var taskConfig = TaskConfig.LoadFromArgs(args);

try
{
    config.Validate();
    transportConfig.Validate();
    sessionConfig.Validate();
    authConfig.Validate();
    trackingConfig.Validate();
    taskConfig.Validate();
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
if (taskConfig.Enabled)
    Console.Error.WriteLine($"[INFO] Async tasks enabled (max concurrent: {taskConfig.MaxConcurrentTasks})");
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

IHost CreateStdioHost()
{
    var builder = Host.CreateApplicationBuilder(args);

    builder.Logging.ClearProviders();
    builder.Logging.AddConsole(options => { options.LogToStandardErrorThreshold = LogLevel.Trace; });

    builder.Services.AddSingleton(transportConfig);
    builder.Services.AddSingleton(sessionConfig);
    builder.Services.AddSingleton(authConfig);
    builder.Services.AddSingleton(authConfig.ApiKey);
    builder.Services.AddSingleton(authConfig.Jwt);
    builder.Services.AddSingleton(trackingConfig);
    builder.Services.AddSingleton(taskConfig);
    builder.Services.AddSingleton<DocumentSessionManager>();
    builder.Services.AddSingleton<TempFileManager>();
    builder.Services.AddHostedService(sp => sp.GetRequiredService<TempFileManager>());
    builder.Services.AddSingleton<ISessionIdentityAccessor, StdioSessionIdentityAccessor>();
    RegisterTaskServices(builder.Services);

    builder.Services.AddMcpServer()
        .WithStdioServerTransport()
        .WithFilteredTools(config, sessionConfig);

    return builder.Build();
}

IHost CreateSseHost(TransportConfig transport)
{
    var builder = CreateWebAppBuilder(transport);
    builder.Services.AddMcpServer().WithFilteredTools(config, sessionConfig);
    var app = builder.Build();

    LogServerStartup($"SSE server listening on http://{transport.Host}:{transport.Port}");
    ConfigureOriginMiddleware(app);
    ConfigureAuthMiddleware(app);
    ConfigureTrackingMiddleware(app);
    MapHealthEndpoints(app);
    app.MapMcp("/mcp");

    return app;
}

IHost CreateWebSocketHost(TransportConfig transport)
{
    var builder = CreateWebAppBuilder(transport);
    var app = builder.Build();

    LogServerStartup($"WebSocket server listening on ws://{transport.Host}:{transport.Port}");
    ConfigureOriginMiddleware(app);
    ConfigureAuthMiddleware(app);
    ConfigureTrackingMiddleware(app);

    app.UseWebSockets();
    MapHealthEndpoints(app);
    ConfigureWebSocketEndpoint(app);

    return app;
}

WebApplicationBuilder CreateWebAppBuilder(TransportConfig transport)
{
    var builder = WebApplication.CreateBuilder(args);
    ConfigureKestrel(builder, transport);
    ConfigureLogging(builder);
    RegisterCoreServices(builder);
    RegisterAuthServices(builder);
    return builder;
}

void ConfigureKestrel(WebApplicationBuilder builder, TransportConfig transport)
{
    builder.WebHost.ConfigureKestrel(options =>
    {
        if (transport.Host == "localhost")
            options.ListenLocalhost(transport.Port);
        else if (transport.Host == "0.0.0.0" || transport.Host == "*")
            options.ListenAnyIP(transport.Port);
        else
            options.Listen(IPAddress.Parse(transport.Host), transport.Port);
    });
}

void ConfigureLogging(WebApplicationBuilder builder)
{
    builder.Logging.ClearProviders();
    builder.Logging.AddConsole(options => { options.LogToStandardErrorThreshold = LogLevel.Trace; });
}

void RegisterCoreServices(WebApplicationBuilder builder)
{
    builder.Services.AddSingleton(transportConfig);
    builder.Services.AddSingleton(sessionConfig);
    builder.Services.AddSingleton(authConfig);
    builder.Services.AddSingleton(authConfig.ApiKey);
    builder.Services.AddSingleton(authConfig.Jwt);
    builder.Services.AddSingleton(trackingConfig);
    builder.Services.AddSingleton(taskConfig);
    builder.Services.AddSingleton<DocumentSessionManager>();
    builder.Services.AddSingleton<TempFileManager>();
    builder.Services.AddHostedService(sp => sp.GetRequiredService<TempFileManager>());
    builder.Services.AddHttpContextAccessor();
    builder.Services.AddSingleton<ISessionIdentityAccessor, HttpContextSessionIdentityAccessor>();
    builder.Services.AddHttpClient();
    RegisterTaskServices(builder.Services);
}

void RegisterTaskServices(IServiceCollection services)
{
    if (!taskConfig.Enabled) return;

    services.AddSingleton<TaskStore>();
    services.AddSingleton<TaskExecutor>();
    services.AddHostedService<TaskCleanupService>();
}

void RegisterAuthServices(WebApplicationBuilder builder)
{
    if (authConfig.ApiKey.Enabled)
        builder.Services.AddSingleton<ApiKeyAuthenticationMiddleware>();
    if (authConfig.Jwt.Enabled)
        builder.Services.AddSingleton<JwtAuthenticationMiddleware>();
}

void LogServerStartup(string message)
{
    Console.Error.WriteLine($"[INFO] {message}");
}

void ConfigureOriginMiddleware(WebApplication app)
{
    if (originConfig.Enabled)
    {
        Console.Error.WriteLine($"[INFO] Origin validation enabled (localhost: {originConfig.AllowLocalhost})");
        app.UseMiddleware<OriginValidationMiddleware>(originConfig);
    }
}

void ConfigureAuthMiddleware(WebApplication app)
{
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
}

void ConfigureTrackingMiddleware(WebApplication app)
{
    if (trackingConfig.LogEnabled || trackingConfig.WebhookEnabled || trackingConfig.MetricsEnabled)
    {
        Console.Error.WriteLine("[INFO] Tracking middleware enabled");
        app.UseMiddleware<TrackingMiddleware>();
    }
}

void MapHealthEndpoints(WebApplication app)
{
    app.MapGet("/health", () => Results.Ok(new { status = "healthy" }));
    app.MapGet("/ready", () => Results.Ok(new { status = "ready" }));
}

void ConfigureWebSocketEndpoint(WebApplication app)
{
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
            var groupId = context.Items["GroupId"]?.ToString();
            var userId = context.Items["UserId"]?.ToString();
            var webSocket = await context.WebSockets.AcceptWebSocketAsync();
            await handler.HandleConnectionAsync(webSocket, context.RequestAborted, groupId, userId);
        }
        else
        {
            context.Response.StatusCode = 400;
        }
    });
}
