using System.ComponentModel;
using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint animations (add, edit, delete, get).
/// </summary>
[McpServerToolType]
public class PptAnimationTool
{
    /// <summary>
    ///     Handler registry for animation operations.
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
    ///     Initializes a new instance of the <see cref="PptAnimationTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public PptAnimationTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry =
            HandlerRegistry<Presentation>.CreateFromNamespace("AsposeMcpServer.Handlers.PowerPoint.Animation");
    }

    /// <summary>
    ///     Executes a PowerPoint animation operation (add, edit, delete, get).
    /// </summary>
    /// <param name="operation">The operation to perform: add, edit, delete, get.</param>
    /// <param name="slideIndex">Slide index (0-based).</param>
    /// <param name="path">Presentation file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="shapeIndex">Shape index (0-based, required for add/edit, optional for delete).</param>
    /// <param name="animationIndex">Animation index (0-based, optional for edit/delete, targets specific animation).</param>
    /// <param name="effectType">Animation effect type (e.g., Fade, Fly, Appear, Bounce, Zoom, Wipe, Split, etc.).</param>
    /// <param name="effectSubtype">Animation effect subtype for direction/style.</param>
    /// <param name="triggerType">Trigger type (OnClick, AfterPrevious, WithPrevious).</param>
    /// <param name="duration">Animation duration in seconds.</param>
    /// <param name="delay">Animation delay in seconds.</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "ppt_animation")]
    [Description(@"Manage PowerPoint animations. Supports 4 operations: add, edit, delete, get.

Usage examples:
- Add animation: ppt_animation(operation='add', path='presentation.pptx', slideIndex=0, shapeIndex=0, effectType='Fade')
- Edit animation: ppt_animation(operation='edit', path='presentation.pptx', slideIndex=0, shapeIndex=0, animationIndex=0, effectType='Fly')
- Delete animation: ppt_animation(operation='delete', path='presentation.pptx', slideIndex=0, shapeIndex=0)
- Get animations: ppt_animation(operation='get', path='presentation.pptx', slideIndex=0)")]
    public string Execute(
        [Description("Operation: add, edit, delete, get")]
        string operation,
        [Description("Slide index (0-based)")] int slideIndex,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Shape index (0-based, required for add/edit, optional for delete)")]
        int? shapeIndex = null,
        [Description("Animation index (0-based, optional for edit/delete, targets specific animation)")]
        int? animationIndex = null,
        [Description("Animation effect type (e.g., Fade, Fly, Appear, Bounce, Zoom, Wipe, Split, etc.)")]
        string? effectType = null,
        [Description(
            "Animation effect subtype for direction/style (e.g., FromBottom, FromLeft, FromRight, FromTop, Horizontal, Vertical)")]
        string? effectSubtype = null,
        [Description("Trigger type (OnClick, AfterPrevious, WithPrevious)")]
        string? triggerType = null,
        [Description("Animation duration in seconds")]
        float? duration = null,
        [Description("Animation delay in seconds")]
        float? delay = null)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, slideIndex, shapeIndex, animationIndex,
            effectType, effectSubtype, triggerType, duration, delay);

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
    ///     Builds OperationParameters from method parameters.
    /// </summary>
    private static OperationParameters BuildParameters(
        string operation,
        int slideIndex,
        int? shapeIndex,
        int? animationIndex,
        string? effectType,
        string? effectSubtype,
        string? triggerType,
        float? duration,
        float? delay)
    {
        var parameters = new OperationParameters();
        parameters.Set("slideIndex", slideIndex);

        switch (operation.ToLowerInvariant())
        {
            case "add":
                if (shapeIndex.HasValue) parameters.Set("shapeIndex", shapeIndex.Value);
                if (effectType != null) parameters.Set("effectType", effectType);
                if (effectSubtype != null) parameters.Set("effectSubtype", effectSubtype);
                if (triggerType != null) parameters.Set("triggerType", triggerType);
                break;

            case "edit":
                if (shapeIndex.HasValue) parameters.Set("shapeIndex", shapeIndex.Value);
                if (animationIndex.HasValue) parameters.Set("animationIndex", animationIndex.Value);
                if (effectType != null) parameters.Set("effectType", effectType);
                if (effectSubtype != null) parameters.Set("effectSubtype", effectSubtype);
                if (triggerType != null) parameters.Set("triggerType", triggerType);
                if (duration.HasValue) parameters.Set("duration", duration.Value);
                if (delay.HasValue) parameters.Set("delay", delay.Value);
                break;

            case "delete":
                if (shapeIndex.HasValue) parameters.Set("shapeIndex", shapeIndex.Value);
                if (animationIndex.HasValue) parameters.Set("animationIndex", animationIndex.Value);
                break;

            case "get":
                if (shapeIndex.HasValue) parameters.Set("shapeIndex", shapeIndex.Value);
                break;
        }

        return parameters;
    }
}
