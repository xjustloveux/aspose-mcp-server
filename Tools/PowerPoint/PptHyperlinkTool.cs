using System.ComponentModel;
using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Handlers.PowerPoint.Hyperlink;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint hyperlinks (add, edit, delete, get)
/// </summary>
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
        _handlerRegistry = PptHyperlinkHandlerRegistry.Create();
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
    [McpServerTool(Name = "ppt_hyperlink")]
    [Description(@"Manage PowerPoint hyperlinks. Supports 4 operations: add, edit, delete, get.

Usage examples:
- Add hyperlink (URL, shape-level): ppt_hyperlink(operation='add', path='presentation.pptx', slideIndex=0, text='Click here', url='https://example.com')
- Add hyperlink (URL, text-level): ppt_hyperlink(operation='add', path='presentation.pptx', slideIndex=0, text='Please click here for more info', linkText='here', url='https://example.com')
- Add hyperlink (internal): ppt_hyperlink(operation='add', path='presentation.pptx', slideIndex=0, text='Go to slide 5', slideTargetIndex=4)
- Edit hyperlink: ppt_hyperlink(operation='edit', path='presentation.pptx', slideIndex=0, shapeIndex=0, url='https://newurl.com')
- Delete hyperlink: ppt_hyperlink(operation='delete', path='presentation.pptx', slideIndex=0, shapeIndex=0)
- Get hyperlinks: ppt_hyperlink(operation='get', path='presentation.pptx', slideIndex=0)")]
    public string Execute(
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

        if (operation.ToLowerInvariant() == "get")
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
        var parameters = new OperationParameters();

        switch (operation.ToLowerInvariant())
        {
            case "add":
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
                break;

            case "edit":
                if (slideIndex.HasValue) parameters.Set("slideIndex", slideIndex.Value);
                if (shapeIndex.HasValue) parameters.Set("shapeIndex", shapeIndex.Value);
                if (url != null) parameters.Set("url", url);
                if (slideTargetIndex.HasValue) parameters.Set("slideTargetIndex", slideTargetIndex.Value);
                if (removeHyperlink) parameters.Set("removeHyperlink", removeHyperlink);
                break;

            case "delete":
                if (slideIndex.HasValue) parameters.Set("slideIndex", slideIndex.Value);
                if (shapeIndex.HasValue) parameters.Set("shapeIndex", shapeIndex.Value);
                break;

            case "get":
                if (slideIndex.HasValue) parameters.Set("slideIndex", slideIndex.Value);
                break;
        }

        return parameters;
    }
}
