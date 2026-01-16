using System.ComponentModel;
using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for PowerPoint data operations (get statistics, get content, get slide details)
///     Merges: PptGetStatisticsTool, PptGetContentTool, PptGetSlideDetailsTool
/// </summary>
[McpServerToolType]
public class PptDataOperationsTool
{
    /// <summary>
    ///     Handler registry for data operations.
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
    ///     Initializes a new instance of the <see cref="PptDataOperationsTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public PptDataOperationsTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry =
            HandlerRegistry<Presentation>.CreateFromNamespace("AsposeMcpServer.Handlers.PowerPoint.DataOperations");
    }

    /// <summary>
    ///     Executes a PowerPoint data operation (get_statistics, get_content, get_slide_details).
    /// </summary>
    /// <param name="operation">The operation to perform: get_statistics, get_content, get_slide_details.</param>
    /// <param name="path">Presentation file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="slideIndex">Slide index (0-based, required for get_slide_details).</param>
    /// <param name="includeThumbnail">Include Base64 encoded thumbnail image (optional for get_slide_details, default false).</param>
    /// <returns>A JSON string containing the requested data (statistics, content, or slide details).</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "ppt_data_operations")]
    [Description(@"PowerPoint data operations. Supports 3 operations: get_statistics, get_content, get_slide_details.

Usage examples:
- Get statistics: ppt_data_operations(operation='get_statistics', path='presentation.pptx')
- Get content: ppt_data_operations(operation='get_content', path='presentation.pptx')
- Get slide details: ppt_data_operations(operation='get_slide_details', path='presentation.pptx', slideIndex=0)
- Get slide details with thumbnail: ppt_data_operations(operation='get_slide_details', path='presentation.pptx', slideIndex=0, includeThumbnail=true)")]
    public string Execute( // NOSONAR S107 - MCP protocol requires multiple parameters
        [Description("Operation: get_statistics, get_content, get_slide_details")]
        string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Slide index (0-based, required for get_slide_details)")]
        int? slideIndex = null,
        [Description("Include Base64 encoded thumbnail image (optional for get_slide_details, default false)")]
        bool includeThumbnail = false)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, slideIndex, includeThumbnail);

        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<Presentation>
        {
            Document = ctx.Document,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = path
        };

        return handler.Execute(operationContext, parameters);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters( // NOSONAR S107
        string operation,
        int? slideIndex,
        bool includeThumbnail)
    {
        return operation.ToLowerInvariant() switch
        {
            "get_slide_details" => BuildGetSlideDetailsParameters(slideIndex, includeThumbnail),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the get_slide_details operation.
    /// </summary>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="includeThumbnail">Whether to include a Base64 encoded thumbnail.</param>
    /// <returns>OperationParameters configured for getting slide details.</returns>
    private static OperationParameters BuildGetSlideDetailsParameters(int? slideIndex, bool includeThumbnail)
    {
        var parameters = new OperationParameters();
        if (slideIndex.HasValue) parameters.Set("slideIndex", slideIndex.Value);
        parameters.Set("includeThumbnail", includeThumbnail);
        return parameters;
    }
}
