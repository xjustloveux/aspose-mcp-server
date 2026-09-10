using System.Diagnostics.CodeAnalysis;
using System.Net;
using AsposeMcpServer.Core.Cleanup;
using AsposeMcpServer.Core.Extension;
using AsposeMcpServer.Core.Security;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Core.Tracking;
using AsposeMcpServer.Core.Transport;
using AsposeMcpServer.Errors;
using AsposeMcpServer.Helpers;
using ModelContextProtocol;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Core;

/// <summary>
///     Factory for creating and configuring host instances based on transport mode.
/// </summary>
internal static class HostFactory
{
    /// <summary>
    ///     Server name for MCP protocol identification.
    /// </summary>
    private const string ServerName = "AsposeMcpServer";

    /// <summary>
    ///     Server description for MCP protocol identification.
    /// </summary>
    private const string ServerDescription =
        "MCP server for document processing with Aspose libraries. " +
        "Supports Word, Excel, PowerPoint, PDF, Email, OCR, and BarCode operations.";

    /// <summary>
    ///     Server website URL for MCP protocol identification.
    /// </summary>
    [SuppressMessage("SonarAnalyzer.CSharp", "S1075",
        Justification = "Static project website URL for MCP protocol identification")]
    private const string ServerWebsiteUrl = "https://xjustloveux.github.io/aspose-mcp-server";

    /// <summary>
    ///     Creates an appropriate host based on the transport configuration.
    /// </summary>
    /// <param name="args">Command line arguments.</param>
    /// <param name="bundle">Configuration bundle containing all host settings.</param>
    /// <returns>The configured host instance.</returns>
    /// <exception cref="ArgumentException">Thrown when transport mode is unknown.</exception>
    public static IHost CreateHost(string[] args, HostConfigBundle bundle)
    {
        return bundle.TransportConfig.Mode switch
        {
            TransportMode.Stdio => CreateStdioHost(args, bundle),
            TransportMode.Http => CreateHttpHost(args, bundle),
            TransportMode.WebSocket => CreateWebSocketHost(args, bundle),
            _ => throw new ArgumentException($"Unknown transport mode: {bundle.TransportConfig.Mode}")
        };
    }

    /// <summary>
    ///     Creates a host configured for stdio transport mode.
    /// </summary>
    /// <param name="args">Command line arguments.</param>
    /// <param name="bundle">Configuration bundle containing all host settings.</param>
    /// <returns>The configured stdio host instance.</returns>
    private static IHost CreateStdioHost(string[] args, HostConfigBundle bundle)
    {
        var builder = Host.CreateApplicationBuilder(args);

        ConfigureLogging(builder.Logging);
        RegisterCoreServices(builder.Services, bundle.TransportConfig, bundle.SessionConfig, bundle.AuthConfig,
            bundle.TrackingConfig, bundle.ExtensionConfig);
        builder.Services.AddSingleton(bundle.ServerConfig);
        builder.Services.AddSingleton<ISessionIdentityAccessor, StdioSessionIdentityAccessor>();

        builder.Services.AddMcpServer(ConfigureServerOptions)
            .WithStdioServerTransport()
            .WithFilteredToolsAndSchemas(bundle.ServerConfig, bundle.SessionConfig, bundle.ExtensionConfig);

        return builder.Build();
    }

    /// <summary>
    ///     Creates a host configured for Streamable HTTP transport mode (MCP 2025-03-26+).
    /// </summary>
    /// <param name="args">Command line arguments.</param>
    /// <param name="bundle">Configuration bundle containing all host settings.</param>
    /// <returns>The configured HTTP host instance.</returns>
    private static IHost CreateHttpHost(string[] args, HostConfigBundle bundle)
    {
        var builder = CreateWebAppBuilder(args, bundle.TransportConfig, bundle.SessionConfig, bundle.AuthConfig,
            bundle.TrackingConfig, bundle.ExtensionConfig);
        builder.Services.AddSingleton(bundle.ServerConfig);
        builder.Services.AddMcpServer(ConfigureServerOptions)
            .WithHttpTransport()
            .WithFilteredToolsAndSchemas(bundle.ServerConfig, bundle.SessionConfig, bundle.ExtensionConfig);
        var app = builder.Build();

        LogServerStartup(
            $"HTTP server listening on http://{bundle.TransportConfig.Host}:{bundle.TransportConfig.Port}/mcp");
        WarnIfReachableWithoutAuthentication(bundle);
        ConfigureMiddleware(app, bundle.AuthConfig, bundle.TrackingConfig, bundle.OriginConfig);
        MapHealthEndpoints(app);
        app.MapMcp("/mcp");

        return app;
    }

    /// <summary>
    ///     Creates a host configured for WebSocket transport mode.
    /// </summary>
    /// <param name="args">Command line arguments.</param>
    /// <param name="bundle">Configuration bundle containing all host settings.</param>
    /// <returns>The configured WebSocket host instance.</returns>
    private static IHost CreateWebSocketHost(string[] args, HostConfigBundle bundle)
    {
        var builder = CreateWebAppBuilder(args, bundle.TransportConfig, bundle.SessionConfig, bundle.AuthConfig,
            bundle.TrackingConfig, bundle.ExtensionConfig);
        builder.Services.AddSingleton(bundle.ServerConfig);
        builder.Services.AddMcpServer(ConfigureServerOptions)
            .WithFilteredToolsAndSchemas(bundle.ServerConfig, bundle.SessionConfig, bundle.ExtensionConfig);
        var app = builder.Build();

        LogServerStartup(
            $"WebSocket server listening on ws://{bundle.TransportConfig.Host}:{bundle.TransportConfig.Port}/mcp");
        WarnIfReachableWithoutAuthentication(bundle);
        ConfigureMiddleware(app, bundle.AuthConfig, bundle.TrackingConfig, bundle.OriginConfig);

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
    /// <param name="extensionConfig">Extension configuration.</param>
    /// <returns>The configured web application builder.</returns>
    private static WebApplicationBuilder CreateWebAppBuilder(
        string[] args,
        TransportConfig transportConfig,
        SessionConfig sessionConfig,
        AuthConfig authConfig,
        TrackingConfig trackingConfig,
        ExtensionConfig extensionConfig)
    {
        var builder = WebApplication.CreateBuilder(args);
        ConfigureKestrel(builder, transportConfig);
        ConfigureLogging(builder.Logging);
        RegisterCoreServices(builder.Services, transportConfig, sessionConfig, authConfig, trackingConfig,
            extensionConfig);
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
    ///     Creates a filter that preserves designed, user-facing exception messages in tool error
    ///     responses (the MCP SDK would otherwise replace them with a generic message) while keeping
    ///     unexpected exception text off the wire: raw BCL/Aspose messages can carry absolute
    ///     file-system paths and internals, so they are replaced by fixed sentinels.
    /// </summary>
    /// <returns>A filter that catches exceptions and returns them as sanitized error results.</returns>
    internal static McpRequestFilter<CallToolRequestParams, CallToolResult> CreateErrorDetailFilter()
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
                    Content = [new TextContentBlock { Text = GetUserFacingErrorText(ex) }]
                };
            }
        };
    }

    /// <summary>
    ///     Maps an exception escaping a tool call to the text returned to the MCP caller.
    ///     Exception types the codebase throws deliberately (parameter validation, addressing,
    ///     session lookup, translator sentinels) keep their in-repo authored message; anything else
    ///     is replaced by a fixed sentinel so raw paths / internals never reach the wire, with the
    ///     original exception written to stderr for server-side diagnosis.
    /// </summary>
    /// <param name="ex">The exception thrown by the tool call.</param>
    /// <returns>The sanitized user-facing error text.</returns>
    private static string GetUserFacingErrorText(Exception ex)
    {
        switch (ex)
        {
            // Designed user-facing channels whose messages are authored in this repository.
            // ObjectDisposedException (session closed) derives from InvalidOperationException.
            case ArgumentException:
            case KeyNotFoundException:
            case NotSupportedException:
            case InvalidOperationException:
            case InvalidCastException:
            case McpException:
                return ex.Message;

            // BCL file errors embed absolute paths ("Could not find file 'C:\...'"); handlers that
            // check existence themselves already throw this exact fixed sentinel.
            case FileNotFoundException:
            case DirectoryNotFoundException:
                return "The specified file was not found.";

            case UnauthorizedAccessException:
                // Sentinels authored by the error translators pass through; raw BCL access-denied
                // text carries the full path. The output-directory message is recognised by shape
                // because it embeds a caller-supplied basename.
                return ex.Message == ErrorMessageBuilder.InvalidPassword()
                       || ErrorMessageBuilder.IsOutputDirectoryNotWritable(ex.Message)
                    ? ex.Message
                    : "Access to the file was denied.";

            // A write that failed on the caller's output directory, reported by a translator. Raw
            // IO text is still replaced below, so only the authored form survives.
            case IOException when ErrorMessageBuilder.IsOutputDirectoryNotWritable(ex.Message):
                return ex.Message;

            default:
                Console.Error.WriteLine($"[WARN] Unhandled tool exception replaced by sentinel: {ex}");
                return ErrorMessageBuilder.ProcessingFailed();
        }
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
    /// <param name="extensionConfig">Extension configuration.</param>
    private static void RegisterCoreServices(
        IServiceCollection services,
        TransportConfig transportConfig,
        SessionConfig sessionConfig,
        AuthConfig authConfig,
        TrackingConfig trackingConfig,
        ExtensionConfig extensionConfig)
    {
        services.AddSingleton(transportConfig);
        services.AddSingleton(sessionConfig);
        services.AddSingleton(authConfig);
        services.AddSingleton(authConfig.ApiKey);
        services.AddSingleton(authConfig.Jwt);
        services.AddSingleton(trackingConfig);
        services.AddSingleton(extensionConfig);
        services.AddSingleton<DocumentSessionManager>();
        services.AddSingleton<TempFileManager>();
        services.AddHostedService(sp => sp.GetRequiredService<TempFileManager>());
        services.AddHostedService<SessionLifetimeService>();
        services.AddHostedService<CleanupDebtService>();

        services.AddSingleton<SnapshotManager>();
        services.AddSingleton<ExtensionManager>();
        services.AddSingleton<ExtensionSessionBridge>();
        services.AddHostedService(sp => sp.GetRequiredService<SnapshotManager>());
        services.AddHostedService(sp => sp.GetRequiredService<ExtensionManager>());
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
    ///     Says, loudly, when a network transport is bound to every interface with no
    ///     authentication enabled.
    /// </summary>
    /// <param name="bundle">The host's configuration.</param>
    /// <remarks>
    ///     The container image sets <c>ASPOSE_HOST=0.0.0.0</c> and exposes the port, and both
    ///     authentication schemes are opt-in, so the combination is one an operator reaches by
    ///     following the deployment guide (R21-DEP01). Only an operator who also chooses a network
    ///     transport and publishes the port makes it reachable, which is why this is a warning
    ///     and not a refusal: behind a reverse proxy or inside a private network it is a
    ///     supported shape. The decision is the warning; a stricter policy would replace this
    ///     method with a refusal and an explicit override flag.
    /// </remarks>
    private static void WarnIfReachableWithoutAuthentication(HostConfigBundle bundle)
    {
        var host = bundle.TransportConfig.Host;
        var everyInterface = host is "0.0.0.0" or "::" or "*" or "+" or "[::]";
        var authenticated = bundle.AuthConfig.ApiKey.Enabled || bundle.AuthConfig.Jwt.Enabled;

        if (!everyInterface || authenticated) return;

        Console.Error.WriteLine(
            "[WARN] This server is listening on every network interface with no authentication "
            + "enabled. Anyone who can reach the port can use every tool. Enable API key or JWT "
            + "authentication, bind to localhost, or place the server behind an authenticating "
            + "proxy before exposing it.");
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

        // After authentication, so an unauthenticated peer is told 401 rather than 400, and before
        // the MCP endpoint, so the SDK never binds a body this server would not (R20-RES02).
        app.UseMiddleware<PayloadShapeMiddleware>();
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
        var (executablePath, prefixArguments) = ChildProcessArguments.ResolveHostCommand();
        var childArguments = prefixArguments.Concat(ChildProcessArguments.BuildChildArguments(args)).ToList();

        var handler = new WebSocketConnectionHandler(executablePath, childArguments,
            app.Services.GetService<ILoggerFactory>());

        app.Map("/mcp", async context =>
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

    /// <summary>
    ///     Configures MCP server options with server identification information.
    /// </summary>
    /// <param name="options">The MCP server options to configure.</param>
    private static void ConfigureServerOptions(McpServerOptions options)
    {
        options.ServerInfo = new Implementation
        {
            Name = ServerName,
            Version = VersionHelper.GetVersion(),
            Description = ServerDescription,
            WebsiteUrl = ServerWebsiteUrl
        };
        options.Filters.Request.CallToolFilters.Add(CreateErrorDetailFilter());
    }

    /// <summary>
    ///     Encapsulates all configuration objects needed for host creation.
    /// </summary>
    /// <param name="ServerConfig">Server configuration.</param>
    /// <param name="TransportConfig">Transport configuration.</param>
    /// <param name="SessionConfig">Session configuration.</param>
    /// <param name="AuthConfig">Authentication configuration.</param>
    /// <param name="TrackingConfig">Tracking configuration.</param>
    /// <param name="OriginConfig">Origin validation configuration.</param>
    /// <param name="ExtensionConfig">Extension configuration.</param>
    internal sealed record HostConfigBundle(
        ServerConfig ServerConfig,
        TransportConfig TransportConfig,
        SessionConfig SessionConfig,
        AuthConfig AuthConfig,
        TrackingConfig TrackingConfig,
        OriginValidationConfig OriginConfig,
        ExtensionConfig ExtensionConfig);
}
