using System.ComponentModel;
using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint transitions (set, get, delete)
///     Merges: PptSetTransitionTool, PptGetTransitionTool, PptDeleteTransitionTool
/// </summary>
[McpServerToolType]
public class PptTransitionTool
{
    /// <summary>
    ///     Handler registry for transition operations.
    /// </summary>
    private readonly HandlerRegistry<Presentation> _handlerRegistry;

    /// <summary>
    ///     Identity accessor for session isolation.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document session handling.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PptTransitionTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory editing.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public PptTransitionTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry =
            HandlerRegistry<Presentation>.CreateFromNamespace("AsposeMcpServer.Handlers.PowerPoint.Transition");
    }

    /// <summary>
    ///     Executes a PowerPoint transition operation (set, get, delete).
    /// </summary>
    /// <param name="operation">The operation to perform: set, get, delete.</param>
    /// <param name="path">Presentation file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (optional, for set/delete operations, defaults to input path).</param>
    /// <param name="slideIndex">Slide index (0-based).</param>
    /// <param name="transitionType">
    ///     Transition type: all TransitionType enum values supported (Fade, Push, Wipe, Split,
    ///     Random, Circle, Plus, Diamond, etc., required for set).
    /// </param>
    /// <param name="advanceAfterSeconds">
    ///     Seconds before auto-advancing to next slide (optional, for set, default: no
    ///     auto-advance).
    /// </param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "ppt_transition")]
    [Description(@"Manage PowerPoint transitions. Supports 3 operations: set, get, delete.

Transition types: Fade, Push, Wipe, Split, Random, Circle, Plus, Diamond, Comb, Cover, Cut, Dissolve, Zoom, and more (all TransitionType enum values supported).

Usage examples:
- Set transition: ppt_transition(operation='set', path='presentation.pptx', slideIndex=0, transitionType='Fade', advanceAfterSeconds=1.5)
- Get transition: ppt_transition(operation='get', path='presentation.pptx', slideIndex=0)
- Delete transition: ppt_transition(operation='delete', path='presentation.pptx', slideIndex=0)")]
    public string Execute(
        [Description(@"Operation to perform.
- 'set': Set slide transition (required params: path, slideIndex, transitionType)
- 'get': Get slide transition (required params: path, slideIndex)
- 'delete': Delete slide transition (required params: path, slideIndex)")]
        string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (optional, for set/delete operations, defaults to input path)")]
        string? outputPath = null,
        [Description("Slide index (0-based)")] int slideIndex = 0,
        [Description(
            "Transition type: all TransitionType enum values supported (Fade, Push, Wipe, Split, Random, Circle, Plus, Diamond, etc., required for set)")]
        string? transitionType = null,
        [Description("Seconds before auto-advancing to next slide (optional, for set, default: no auto-advance)")]
        double? advanceAfterSeconds = null)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, slideIndex, transitionType, advanceAfterSeconds);

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
            return result;

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return $"{result}\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters(
        string operation,
        int slideIndex,
        string? transitionType,
        double? advanceAfterSeconds)
    {
        return operation.ToLowerInvariant() switch
        {
            "set" => BuildSetParameters(slideIndex, transitionType, advanceAfterSeconds),
            "get" or "delete" => BuildBaseParameters(slideIndex),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the set transition operation.
    /// </summary>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="transitionType">The transition type (e.g., Fade, Push, Wipe).</param>
    /// <param name="advanceAfterSeconds">Seconds before auto-advancing to next slide.</param>
    /// <returns>OperationParameters configured for setting a transition.</returns>
    private static OperationParameters BuildSetParameters(int slideIndex, string? transitionType,
        double? advanceAfterSeconds)
    {
        var parameters = new OperationParameters();
        parameters.Set("slideIndex", slideIndex);
        if (transitionType != null) parameters.Set("transitionType", transitionType);
        if (advanceAfterSeconds.HasValue) parameters.Set("advanceAfterSeconds", advanceAfterSeconds.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds base parameters containing only the slide index.
    /// </summary>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <returns>OperationParameters with slide index set.</returns>
    private static OperationParameters BuildBaseParameters(int slideIndex)
    {
        var parameters = new OperationParameters();
        parameters.Set("slideIndex", slideIndex);
        return parameters;
    }
}
