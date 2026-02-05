using System.ComponentModel;
using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint watermarks (add_text, add_image, remove, get).
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.PowerPoint.Watermark")]
[McpServerToolType]
public class PptWatermarkTool
{
    /// <summary>
    ///     Handler registry for watermark operations.
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
    ///     Initializes a new instance of the <see cref="PptWatermarkTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public PptWatermarkTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry =
            HandlerRegistry<Presentation>.CreateFromNamespace("AsposeMcpServer.Handlers.PowerPoint.Watermark");
    }

    /// <summary>
    ///     Executes a PowerPoint watermark operation (add_text, add_image, remove, get).
    /// </summary>
    /// <param name="operation">The operation to perform: add_text, add_image, remove, get.</param>
    /// <param name="path">Presentation file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="text">Watermark text (required for add_text).</param>
    /// <param name="imagePath">Image file path (required for add_image).</param>
    /// <param name="fontSize">Font size for text watermark (default: 48).</param>
    /// <param name="fontColor">Font color for text watermark (default: gray).</param>
    /// <param name="opacity">Opacity 0-255 for text watermark (default: 128).</param>
    /// <param name="rotation">Rotation in degrees for text watermark (default: -45).</param>
    /// <param name="width">Width in points for image watermark (default: 200).</param>
    /// <param name="height">Height in points for image watermark (default: 200).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "ppt_watermark",
        Title = "PowerPoint Watermark Operations",
        Destructive = false,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(@"Manage PowerPoint watermarks. Supports 4 operations: add_text, add_image, remove, get.

Usage examples:
- Add text watermark: ppt_watermark(operation='add_text', path='file.pptx', text='CONFIDENTIAL')
- Add image watermark: ppt_watermark(operation='add_image', path='file.pptx', imagePath='logo.png')
- Remove watermarks: ppt_watermark(operation='remove', path='file.pptx')
- Get watermarks: ppt_watermark(operation='get', path='file.pptx')")]
    public object Execute(
        [Description(@"Operation to perform.
- 'add_text': Add text watermark to all slides (required: text)
- 'add_image': Add image watermark to all slides (required: imagePath)
- 'remove': Remove all watermarks
- 'get': List all watermarks")]
        string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Watermark text (required for add_text)")]
        string? text = null,
        [Description("Image file path (required for add_image)")]
        string? imagePath = null,
        [Description("Font size (for add_text, default: 48)")]
        float fontSize = 48,
        [Description("Font color (for add_text, default: gray)")]
        string fontColor = "128,128,128",
        [Description("Opacity 0-255 (for add_text, default: 128)")]
        int opacity = 128,
        [Description("Rotation in degrees (for add_text, default: -45)")]
        float rotation = -45,
        [Description("Width in points (for add_image, default: 200)")]
        float width = 200,
        [Description("Height in points (for add_image, default: 200)")]
        float height = 200)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, text, imagePath, fontSize, fontColor, opacity, rotation, width,
            height);
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

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return ResultHelper.FinalizeResult((dynamic)result, ctx, outputPath);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters(string operation, string? text, string? imagePath,
        float fontSize, string fontColor, int opacity, float rotation, float width, float height)
    {
        var parameters = new OperationParameters();

        return operation.ToLowerInvariant() switch
        {
            "add_text" => BuildAddTextParameters(parameters, text, fontSize, fontColor, opacity, rotation),
            "add_image" => BuildAddImageParameters(parameters, imagePath, width, height),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for the add text watermark operation.
    /// </summary>
    /// <param name="p">The operation parameters to populate.</param>
    /// <param name="text">The watermark text.</param>
    /// <param name="fontSize">The font size in points.</param>
    /// <param name="fontColor">The font color string.</param>
    /// <param name="opacity">The opacity value (0-255).</param>
    /// <param name="rotation">The rotation angle in degrees.</param>
    /// <returns>OperationParameters configured for adding a text watermark.</returns>
    private static OperationParameters BuildAddTextParameters(OperationParameters p, string? text,
        float fontSize, string fontColor, int opacity, float rotation)
    {
        if (text != null) p.Set("text", text);
        p.Set("fontSize", fontSize);
        p.Set("fontColor", fontColor);
        p.Set("opacity", opacity);
        p.Set("rotation", rotation);
        return p;
    }

    /// <summary>
    ///     Builds parameters for the add image watermark operation.
    /// </summary>
    /// <param name="p">The operation parameters to populate.</param>
    /// <param name="imagePath">The image file path.</param>
    /// <param name="width">The width in points.</param>
    /// <param name="height">The height in points.</param>
    /// <returns>OperationParameters configured for adding an image watermark.</returns>
    private static OperationParameters BuildAddImageParameters(OperationParameters p, string? imagePath,
        float width, float height)
    {
        if (imagePath != null) p.Set("imagePath", imagePath);
        p.Set("width", width);
        p.Set("height", height);
        return p;
    }
}
