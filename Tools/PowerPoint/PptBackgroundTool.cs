using System.ComponentModel;
using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint backgrounds (set, get).
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.PowerPoint.Background")]
[McpServerToolType]
public class PptBackgroundTool
{
    /// <summary>
    ///     Handler registry for background operations.
    /// </summary>
    private readonly HandlerRegistry<Presentation> _handlerRegistry;

    /// <summary>
    ///     Identity accessor for session isolation.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document lifecycle management.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PptBackgroundTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public PptBackgroundTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry =
            HandlerRegistry<Presentation>.CreateFromNamespace("AsposeMcpServer.Handlers.PowerPoint.Background");
    }

    /// <summary>
    ///     Executes a PowerPoint background operation (set, get).
    /// </summary>
    /// <param name="operation">The operation to perform: set, get.</param>
    /// <param name="path">Presentation file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="slideIndex">Slide index (0-based, default: 0, ignored if applyToAll is true).</param>
    /// <param name="color">Hex color like #FFAA00 or #80FFAA00 (with alpha).</param>
    /// <param name="imagePath">Background image path.</param>
    /// <param name="applyToAll">Apply background to all slides (default: false).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "ppt_background",
        Title = "PowerPoint Background Operations",
        Destructive = false,
        Idempotent = true,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(@"Manage PowerPoint backgrounds. Supports 2 operations: set, get.

Usage examples:
- Set background color: ppt_background(operation='set', path='presentation.pptx', slideIndex=0, color='#FFFFFF')
- Set background image: ppt_background(operation='set', path='presentation.pptx', slideIndex=0, imagePath='bg.png')
- Apply to all slides: ppt_background(operation='set', path='presentation.pptx', color='#FFFFFF', applyToAll=true)
- Get background: ppt_background(operation='get', path='presentation.pptx', slideIndex=0)")]
    public object Execute(
        [Description("Operation: set, get")] string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Slide index (0-based, default: 0, ignored if applyToAll is true)")]
        int slideIndex = 0,
        [Description("Hex color like #FFAA00 or #80FFAA00 (with alpha)")]
        string? color = null,
        [Description("Background image path")] string? imagePath = null,
        [Description("Apply background to all slides (default: false)")]
        bool applyToAll = false)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, slideIndex, color, imagePath, applyToAll);

        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<Presentation>
        {
            Document = ctx.Document,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = path,
            OutputPath = outputPath
        };

        var result = handler.Execute(operationContext, parameters);

        if (string.Equals(operation, "get", StringComparison.OrdinalIgnoreCase))
            return ResultHelper.FinalizeResult((dynamic)result, ctx, outputPath);

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return ResultHelper.FinalizeResult((dynamic)result, ctx, outputPath);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters(
        string operation,
        int slideIndex,
        string? color,
        string? imagePath,
        bool applyToAll)
    {
        return operation.ToLowerInvariant() switch
        {
            "set" => BuildSetParameters(slideIndex, color, imagePath, applyToAll),
            "get" => BuildGetParameters(slideIndex),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the set background operation.
    /// </summary>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="color">Hex color value for background.</param>
    /// <param name="imagePath">Path to background image.</param>
    /// <param name="applyToAll">Whether to apply background to all slides.</param>
    /// <returns>OperationParameters configured for setting background.</returns>
    private static OperationParameters BuildSetParameters(int slideIndex, string? color, string? imagePath,
        bool applyToAll)
    {
        var parameters = new OperationParameters();
        parameters.Set("slideIndex", slideIndex);
        if (color != null) parameters.Set("color", color);
        if (imagePath != null) parameters.Set("imagePath", imagePath);
        parameters.Set("applyToAll", applyToAll);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the get background operation.
    /// </summary>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <returns>OperationParameters configured for getting background info.</returns>
    private static OperationParameters BuildGetParameters(int slideIndex)
    {
        var parameters = new OperationParameters();
        parameters.Set("slideIndex", slideIndex);
        return parameters;
    }
}
