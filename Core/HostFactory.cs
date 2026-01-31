using System.Net;
using AsposeMcpServer.Core.Security;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Core.Tracking;
using AsposeMcpServer.Core.Transport;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Core;

/// <summary>
///     Factory for creating and configuring host instances based on transport mode.
/// </summary>
internal static class HostFactory
{
    /// <summary>
    ///     Creates an appropriate host based on the transport configuration.
    /// </summary>
    /// <param name="args">Command line arguments.</param>
    /// <param name="config">Server configuration.</param>
    /// <param name="transportConfig">Transport configuration.</param>
    /// <param name="sessionConfig">Session configuration.</param>
    /// <param name="authConfig">Authentication configuration.</param>
    /// <param name="trackingConfig">Tracking configuration.</param>
    /// <param name="originConfig">Origin validation configuration.</param>
    /// <returns>The configured host instance.</returns>
    /// <exception cref="ArgumentException">Thrown when transport mode is unknown.</exception>
    public static IHost CreateHost(
        string[] args,
        ServerConfig config,
        TransportConfig transportConfig,
        SessionConfig sessionConfig,
        AuthConfig authConfig,
        TrackingConfig trackingConfig,
        OriginValidationConfig originConfig)
    {
        return transportConfig.Mode switch
        {
            TransportMode.Stdio => CreateStdioHost(args, transportConfig, sessionConfig, authConfig, trackingConfig,
                config),
            TransportMode.Http => CreateHttpHost(args, transportConfig, sessionConfig, authConfig, trackingConfig,
                originConfig, config),
            TransportMode.WebSocket => CreateWebSocketHost(args, transportConfig, sessionConfig, authConfig,
                trackingConfig, originConfig, config),
            _ => throw new ArgumentException($"Unknown transport mode: {transportConfig.Mode}")
        };
    }

    /// <summary>
    ///     Creates a host configured for stdio transport mode.
    /// </summary>
    /// <param name="args">Command line arguments.</param>
    /// <param name="transportConfig">Transport configuration.</param>
    /// <param name="sessionConfig">Session configuration.</param>
    /// <param name="authConfig">Authentication configuration.</param>
    /// <param name="trackingConfig">Tracking configuration.</param>
    /// <param name="config">Server configuration.</param>
    /// <returns>The configured stdio host instance.</returns>
    private static IHost CreateStdioHost(
        string[] args,
        TransportConfig transportConfig,
        SessionConfig sessionConfig,
        AuthConfig authConfig,
        TrackingConfig trackingConfig,
        ServerConfig config)
    {
        var builder = Host.CreateApplicationBuilder(args);

        ConfigureLogging(builder.Logging);
        RegisterCoreServices(builder.Services, transportConfig, sessionConfig, authConfig, trackingConfig);
        builder.Services.AddSingleton(config);
        builder.Services.AddSingleton<ISessionIdentityAccessor, StdioSessionIdentityAccessor>();

        builder.Services.AddMcpServer()
            .WithStdioServerTransport()
            .WithFilteredToolsAndSchemas(config, sessionConfig)
            .AddCallToolFilter(CreateErrorDetailFilter());

        return builder.Build();
    }

    /// <summary>
    ///     Creates a host configured for Streamable HTTP transport mode (MCP 2025-03-26+).
    /// </summary>
    /// <param name="args">Command line arguments.</param>
    /// <param name="transportConfig">Transport configuration.</param>
    /// <param name="sessionConfig">Session configuration.</param>
    /// <param name="authConfig">Authentication configuration.</param>
    /// <param name="trackingConfig">Tracking configuration.</param>
    /// <param name="originConfig">Origin validation configuration.</param>
    /// <param name="config">Server configuration.</param>
    /// <returns>The configured HTTP host instance.</returns>
    private static IHost CreateHttpHost(
        string[] args,
        TransportConfig transportConfig,
        SessionConfig sessionConfig,
        AuthConfig authConfig,
        TrackingConfig trackingConfig,
        OriginValidationConfig originConfig,
        ServerConfig config)
    {
        var builder = CreateWebAppBuilder(args, transportConfig, sessionConfig, authConfig, trackingConfig);
        builder.Services.AddSingleton(config);
        builder.Services.AddMcpServer()
            .WithHttpTransport()
            .WithFilteredToolsAndSchemas(config, sessionConfig)
            .AddCallToolFilter(CreateErrorDetailFilter());
        var app = builder.Build();

        LogServerStartup($"HTTP server listening on http://{transportConfig.Host}:{transportConfig.Port}");
        ConfigureMiddleware(app, authConfig, trackingConfig, originConfig);
        MapHealthEndpoints(app);
        app.MapMcp("/mcp");

        return app;
    }

    /// <summary>
    ///     Creates a host configured for WebSocket transport mode.
    /// </summary>
    /// <param name="args">Command line arguments.</param>
    /// <param name="transportConfig">Transport configuration.</param>
    /// <param name="sessionConfig">Session configuration.</param>
    /// <param name="authConfig">Authentication configuration.</param>
    /// <param name="trackingConfig">Tracking configuration.</param>
    /// <param name="originConfig">Origin validation configuration.</param>
    /// <param name="config">Server configuration.</param>
    /// <returns>The configured WebSocket host instance.</returns>
    private static IHost CreateWebSocketHost(
        string[] args,
        TransportConfig transportConfig,
        SessionConfig sessionConfig,
        AuthConfig authConfig,
        TrackingConfig trackingConfig,
        OriginValidationConfig originConfig,
        ServerConfig config)
    {
        var builder = CreateWebAppBuilder(args, transportConfig, sessionConfig, authConfig, trackingConfig);
        builder.Services.AddSingleton(config);
        builder.Services.AddMcpServer()
            .WithFilteredToolsAndSchemas(config, sessionConfig)
            .AddCallToolFilter(CreateErrorDetailFilter());
        var app = builder.Build();

        LogServerStartup($"WebSocket server listening on ws://{transportConfig.Host}:{transportConfig.Port}");
        ConfigureMiddleware(app, authConfig, trackingConfig, originConfig);

        app.UseWebSockets();
        MapHealthEndpoints(app);
        ConfigureWebSocketEndpoint(app, args);

        return app;
    }

    /// <summary>
    ///     Creates a web application builder with common configuration for HTTP-based transports.
    /// </summary>
    /// <param name="args">Command line arguments.</param>
    /// <param name="transportConfig">Transport configuration.</param>
    /// <param name="sessionConfig">Session configuration.</param>
    /// <param name="authConfig">Authentication configuration.</param>
    /// <param name="trackingConfig">Tracking configuration.</param>
    /// <returns>The configured web application builder.</returns>
    private static WebApplicationBuilder CreateWebAppBuilder(
        string[] args,
        TransportConfig transportConfig,
        SessionConfig sessionConfig,
        AuthConfig authConfig,
        TrackingConfig trackingConfig)
    {
        var builder = WebApplication.CreateBuilder(args);
        ConfigureKestrel(builder, transportConfig);
        ConfigureLogging(builder.Logging);
        RegisterCoreServices(builder.Services, transportConfig, sessionConfig, authConfig, trackingConfig);
        builder.Services.AddHttpContextAccessor();
        builder.Services.AddSingleton<ISessionIdentityAccessor, HttpContextSessionIdentityAccessor>();
        builder.Services.AddHttpClient();
        RegisterAuthServices(builder.Services, authConfig);
        return builder;
    }

    /// <summary>
    ///     Configures Kestrel web server with the specified transport settings.
    /// </summary>
    /// <param name="builder">The web application builder to configure.</param>
    /// <param name="transport">Transport configuration specifying host and port.</param>
    private static void ConfigureKestrel(WebApplicationBuilder builder, TransportConfig transport)
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

    /// <summary>
    ///     Creates a filter that preserves original exception messages in tool error responses.
    ///     Without this filter, the MCP SDK replaces exception details with a generic error message.
    /// </summary>
    /// <returns>A filter that catches exceptions and returns them as detailed error results.</returns>
    private static McpRequestFilter<CallToolRequestParams, CallToolResult> CreateErrorDetailFilter()
    {
        return next => async (request, cancellationToken) =>
        {
            try
            {
                return await next(request, cancellationToken);
            }
            catch (Exception ex)
            {
                return new CallToolResult
                {
                    IsError = true,
                    Content = [new TextContentBlock { Text = ex.Message }]
                };
            }
        };
    }

    /// <summary>
    ///     Configures logging to output to standard error with trace level threshold.
    /// </summary>
    /// <param name="logging">The logging builder to configure.</param>
    private static void ConfigureLogging(ILoggingBuilder logging)
    {
        logging.ClearProviders();
        logging.AddConsole(options => { options.LogToStandardErrorThreshold = LogLevel.Trace; });
    }

    /// <summary>
    ///     Registers core services required by all transport modes.
    /// </summary>
    /// <param name="services">The service collection to register services into.</param>
    /// <param name="transportConfig">Transport configuration.</param>
    /// <param name="sessionConfig">Session configuration.</param>
    /// <param name="authConfig">Authentication configuration.</param>
    /// <param name="trackingConfig">Tracking configuration.</param>
    private static void RegisterCoreServices(
        IServiceCollection services,
        TransportConfig transportConfig,
        SessionConfig sessionConfig,
        AuthConfig authConfig,
        TrackingConfig trackingConfig)
    {
        services.AddSingleton(transportConfig);
        services.AddSingleton(sessionConfig);
        services.AddSingleton(authConfig);
        services.AddSingleton(authConfig.ApiKey);
        services.AddSingleton(authConfig.Jwt);
        services.AddSingleton(trackingConfig);
        services.AddSingleton<DocumentSessionManager>();
        services.AddSingleton<TempFileManager>();
        services.AddHostedService(sp => sp.GetRequiredService<TempFileManager>());
        services.AddHostedService<SessionLifetimeService>();
    }

    /// <summary>
    ///     Registers authentication middleware services based on configuration.
    /// </summary>
    /// <param name="services">The service collection to register services into.</param>
    /// <param name="authConfig">Authentication configuration specifying enabled auth methods.</param>
    private static void RegisterAuthServices(IServiceCollection services, AuthConfig authConfig)
    {
        if (authConfig.ApiKey.Enabled)
            services.AddSingleton<ApiKeyAuthenticationMiddleware>();
        if (authConfig.Jwt.Enabled)
            services.AddSingleton<JwtAuthenticationMiddleware>();
    }

    /// <summary>
    ///     Logs a server startup message to standard error.
    /// </summary>
    /// <param name="message">The message to log.</param>
    private static void LogServerStartup(string message)
    {
        Console.Error.WriteLine($"[INFO] {message}");
    }

    /// <summary>
    ///     Configures all middleware components for HTTP-based transports.
    /// </summary>
    /// <param name="app">The web application to configure.</param>
    /// <param name="authConfig">Authentication configuration.</param>
    /// <param name="trackingConfig">Tracking configuration.</param>
    /// <param name="originConfig">Origin validation configuration.</param>
    private static void ConfigureMiddleware(
        WebApplication app,
        AuthConfig authConfig,
        TrackingConfig trackingConfig,
        OriginValidationConfig originConfig)
    {
        ConfigureOriginMiddleware(app, originConfig);
        ConfigureAuthMiddleware(app, authConfig);
        ConfigureTrackingMiddleware(app, trackingConfig);
    }

    /// <summary>
    ///     Configures origin validation middleware if enabled.
    /// </summary>
    /// <param name="app">The web application to configure.</param>
    /// <param name="originConfig">Origin validation configuration.</param>
    private static void ConfigureOriginMiddleware(WebApplication app, OriginValidationConfig originConfig)
    {
        if (originConfig.Enabled)
        {
            Console.Error.WriteLine($"[INFO] Origin validation enabled (localhost: {originConfig.AllowLocalhost})");
            app.UseMiddleware<OriginValidationMiddleware>(originConfig);
        }
    }

    /// <summary>
    ///     Configures authentication middleware based on enabled auth methods.
    /// </summary>
    /// <param name="app">The web application to configure.</param>
    /// <param name="authConfig">Authentication configuration.</param>
    private static void ConfigureAuthMiddleware(WebApplication app, AuthConfig authConfig)
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

    /// <summary>
    ///     Configures tracking middleware if any tracking feature is enabled.
    /// </summary>
    /// <param name="app">The web application to configure.</param>
    /// <param name="trackingConfig">Tracking configuration.</param>
    private static void ConfigureTrackingMiddleware(WebApplication app, TrackingConfig trackingConfig)
    {
        if (trackingConfig.LogEnabled || trackingConfig.WebhookEnabled || trackingConfig.MetricsEnabled)
        {
            Console.Error.WriteLine("[INFO] Tracking middleware enabled");
            app.UseMiddleware<TrackingMiddleware>();
        }
    }

    /// <summary>
    ///     Maps health check endpoints for monitoring.
    /// </summary>
    /// <param name="app">The web application to configure.</param>
    private static void MapHealthEndpoints(WebApplication app)
    {
        app.MapGet("/health", () => Microsoft.AspNetCore.Http.Results.Ok(new { status = "healthy" }));
        app.MapGet("/ready", () => Microsoft.AspNetCore.Http.Results.Ok(new { status = "ready" }));
    }

    /// <summary>
    ///     Configures the WebSocket endpoint and connection handler.
    /// </summary>
    /// <param name="app">The web application to configure.</param>
    /// <param name="args">Command line arguments for tool configuration passthrough.</param>
    private static void ConfigureWebSocketEndpoint(WebApplication app, string[] args)
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

        var handler = new WebSocketConnectionHandler(executablePath, toolArgs,
            app.Services.GetService<ILoggerFactory>());

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
}
