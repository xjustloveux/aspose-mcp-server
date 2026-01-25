using System.ComponentModel;
using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint hyperlinks (add, edit, delete, get)
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.PowerPoint.Hyperlink")]
[McpServerToolType]
public class PptHyperlinkTool
{
    /// <summary>
    ///     Handler registry for hyperlink operations.
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
    ///     Initializes a new instance of the <see cref="PptHyperlinkTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public PptHyperlinkTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry =
            HandlerRegistry<Presentation>.CreateFromNamespace("AsposeMcpServer.Handlers.PowerPoint.Hyperlink");
    }

    /// <summary>
    ///     Executes a PowerPoint hyperlink operation (add, edit, delete, get).
    /// </summary>
    /// <param name="operation">The operation to perform: add, edit, delete, get.</param>
    /// <param name="path">Presentation file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="slideIndex">Slide index (0-based).</param>
    /// <param name="shapeIndex">Shape index (0-based, required for edit/delete, optional for add).</param>
    /// <param name="text">Display text (required for add).</param>
    /// <param name="linkText">Specific text to apply hyperlink to (optional, for add).</param>
    /// <param name="url">Hyperlink URL (required for add, optional for edit).</param>
    /// <param name="slideTargetIndex">Target slide index for internal link (0-based, optional, for add/edit).</param>
    /// <param name="removeHyperlink">Remove hyperlink (optional, for edit).</param>
    /// <param name="x">X position (optional, for add, default: 50).</param>
    /// <param name="y">Y position (optional, for add, default: 50).</param>
    /// <param name="width">Width (optional, for add, default: 300).</param>
    /// <param name="height">Height (optional, for add, default: 50).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "ppt_hyperlink",
        Title = "PowerPoint Hyperlink Operations",
        Destructive = true,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(@"Manage PowerPoint hyperlinks. Supports 4 operations: add, edit, delete, get.

Usage examples:
- Add hyperlink (URL, shape-level): ppt_hyperlink(operation='add', path='presentation.pptx', slideIndex=0, text='Click here', url='https://example.com')
- Add hyperlink (URL, text-level): ppt_hyperlink(operation='add', path='presentation.pptx', slideIndex=0, text='Please click here for more info', linkText='here', url='https://example.com')
- Add hyperlink (internal): ppt_hyperlink(operation='add', path='presentation.pptx', slideIndex=0, text='Go to slide 5', slideTargetIndex=4)
- Edit hyperlink: ppt_hyperlink(operation='edit', path='presentation.pptx', slideIndex=0, shapeIndex=0, url='https://newurl.com')
- Delete hyperlink: ppt_hyperlink(operation='delete', path='presentation.pptx', slideIndex=0, shapeIndex=0)
- Get hyperlinks: ppt_hyperlink(operation='get', path='presentation.pptx', slideIndex=0)")]
    public object Execute(
        [Description("Operation: add, edit, delete, get")]
        string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Slide index (0-based)")] int? slideIndex = null,
        [Description("Shape index (0-based, required for edit/delete, optional for add)")]
        int? shapeIndex = null,
        [Description("Display text (required for add)")]
        string? text = null,
        [Description(
            "Specific text to apply hyperlink to (optional, for add). When provided, only this text portion will have the hyperlink.")]
        string? linkText = null,
        [Description("Hyperlink URL (required for add, optional for edit)")]
        string? url = null,
        [Description("Target slide index for internal link (0-based, optional, for add/edit)")]
        int? slideTargetIndex = null,
        [Description("Remove hyperlink (optional, for edit)")]
        bool removeHyperlink = false,
        [Description("X position (optional, for add, default: 50)")]
        float x = 50,
        [Description("Y position (optional, for add, default: 50)")]
        float y = 50,
        [Description("Width (optional, for add, default: 300)")]
        float width = 300,
        [Description("Height (optional, for add, default: 50)")]
        float height = 50)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, slideIndex, shapeIndex, text, linkText,
            url, slideTargetIndex, removeHyperlink, x, y, width, height);

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
        int? slideIndex,
        int? shapeIndex,
        string? text,
        string? linkText,
        string? url,
        int? slideTargetIndex,
        bool removeHyperlink,
        float x,
        float y,
        float width,
        float height)
    {
        return operation.ToLowerInvariant() switch
        {
            "add" => BuildAddParameters(slideIndex, shapeIndex, text, linkText, url, slideTargetIndex, x, y, width,
                height),
            "edit" => BuildEditParameters(slideIndex, shapeIndex, url, slideTargetIndex, removeHyperlink),
            "delete" => BuildDeleteParameters(slideIndex, shapeIndex),
            "get" => BuildGetParameters(slideIndex),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the add hyperlink operation.
    /// </summary>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="shapeIndex">The shape index (0-based).</param>
    /// <param name="text">The display text.</param>
    /// <param name="linkText">The specific text to apply hyperlink to.</param>
    /// <param name="url">The hyperlink URL.</param>
    /// <param name="slideTargetIndex">The target slide index for internal link.</param>
    /// <param name="x">The X position in points.</param>
    /// <param name="y">The Y position in points.</param>
    /// <param name="width">The width in points.</param>
    /// <param name="height">The height in points.</param>
    /// <returns>OperationParameters configured for adding a hyperlink.</returns>
    private static OperationParameters BuildAddParameters(int? slideIndex, int? shapeIndex, string? text,
        string? linkText, string? url, int? slideTargetIndex, float x, float y, float width, float height)
    {
        var parameters = new OperationParameters();
        if (slideIndex.HasValue) parameters.Set("slideIndex", slideIndex.Value);
        if (shapeIndex.HasValue) parameters.Set("shapeIndex", shapeIndex.Value);
        if (text != null) parameters.Set("text", text);
        if (linkText != null) parameters.Set("linkText", linkText);
        if (url != null) parameters.Set("url", url);
        if (slideTargetIndex.HasValue) parameters.Set("slideTargetIndex", slideTargetIndex.Value);
        parameters.Set("x", x);
        parameters.Set("y", y);
        parameters.Set("width", width);
        parameters.Set("height", height);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the edit hyperlink operation.
    /// </summary>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="shapeIndex">The shape index (0-based).</param>
    /// <param name="url">The new hyperlink URL.</param>
    /// <param name="slideTargetIndex">The new target slide index for internal link.</param>
    /// <param name="removeHyperlink">Whether to remove the hyperlink.</param>
    /// <returns>OperationParameters configured for editing a hyperlink.</returns>
    private static OperationParameters BuildEditParameters(int? slideIndex, int? shapeIndex, string? url,
        int? slideTargetIndex, bool removeHyperlink)
    {
        var parameters = new OperationParameters();
        if (slideIndex.HasValue) parameters.Set("slideIndex", slideIndex.Value);
        if (shapeIndex.HasValue) parameters.Set("shapeIndex", shapeIndex.Value);
        if (url != null) parameters.Set("url", url);
        if (slideTargetIndex.HasValue) parameters.Set("slideTargetIndex", slideTargetIndex.Value);
        if (removeHyperlink) parameters.Set("removeHyperlink", removeHyperlink);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the delete hyperlink operation.
    /// </summary>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="shapeIndex">The shape index (0-based).</param>
    /// <returns>OperationParameters configured for deleting a hyperlink.</returns>
    private static OperationParameters BuildDeleteParameters(int? slideIndex, int? shapeIndex)
    {
        var parameters = new OperationParameters();
        if (slideIndex.HasValue) parameters.Set("slideIndex", slideIndex.Value);
        if (shapeIndex.HasValue) parameters.Set("shapeIndex", shapeIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the get hyperlinks operation.
    /// </summary>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <returns>OperationParameters configured for getting hyperlinks.</returns>
    private static OperationParameters BuildGetParameters(int? slideIndex)
    {
        var parameters = new OperationParameters();
        if (slideIndex.HasValue) parameters.Set("slideIndex", slideIndex.Value);
        return parameters;
    }
}
