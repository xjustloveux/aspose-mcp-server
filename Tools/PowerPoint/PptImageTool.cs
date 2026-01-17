using System.ComponentModel;
using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint images.
///     Supports: add, edit, delete, get, export_slides, extract
/// </summary>
[McpServerToolType]
public class PptImageTool
{
    /// <summary>
    ///     Handler registry for image operations.
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
    ///     Initializes a new instance of the <see cref="PptImageTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public PptImageTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry =
            HandlerRegistry<Presentation>.CreateFromNamespace("AsposeMcpServer.Handlers.PowerPoint.Image");
    }

    /// <summary>
    ///     Executes a PowerPoint image operation (add, edit, delete, get, export_slides, extract).
    /// </summary>
    /// <param name="operation">The operation to perform: add, edit, delete, get, export_slides, extract.</param>
    /// <param name="path">Presentation file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="slideIndex">Slide index (0-based, required for add/edit/delete/get).</param>
    /// <param name="imageIndex">Image index on the slide (0-based, required for edit/delete).</param>
    /// <param name="imagePath">Image file path (required for add, optional for edit).</param>
    /// <param name="x">X position in points (optional for add/edit, default: 100).</param>
    /// <param name="y">Y position in points (optional for add/edit, default: 100).</param>
    /// <param name="width">Width in points (optional for add/edit).</param>
    /// <param name="height">Height in points (optional for add/edit).</param>
    /// <param name="jpegQuality">JPEG quality 10-100 (optional for edit, re-encode image as JPEG).</param>
    /// <param name="maxWidth">Maximum width in pixels for resize (optional for edit).</param>
    /// <param name="maxHeight">Maximum height in pixels for resize (optional for edit).</param>
    /// <param name="outputDir">Output directory (required for export_slides/extract).</param>
    /// <param name="format">Image format: png|jpeg (optional for export_slides/extract, default: png).</param>
    /// <param name="scale">Scaling factor (optional for export_slides, default: 1.0).</param>
    /// <param name="slideIndexes">Comma-separated slide indexes to export (optional for export_slides).</param>
    /// <param name="skipDuplicates">Skip duplicate images based on content hash (optional for extract, default: false).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "ppt_image")]
    [Description(@"Manage PowerPoint images. Supports 6 operations: add, edit, delete, get, export_slides, extract.

Usage examples:
- Add image: ppt_image(operation='add', path='presentation.pptx', slideIndex=0, imagePath='image.png', x=100, y=100)
- Edit image: ppt_image(operation='edit', path='presentation.pptx', slideIndex=0, imageIndex=0, width=300, height=200)
- Edit with compression: ppt_image(operation='edit', path='presentation.pptx', slideIndex=0, imageIndex=0, imagePath='new.png', jpegQuality=80, maxWidth=800)
- Delete image: ppt_image(operation='delete', path='presentation.pptx', slideIndex=0, imageIndex=0)
- Get image info: ppt_image(operation='get', path='presentation.pptx', slideIndex=0)
- Export slides as images: ppt_image(operation='export_slides', path='presentation.pptx', outputDir='images/', slideIndexes='0,2,4')
- Extract embedded images: ppt_image(operation='extract', path='presentation.pptx', outputDir='images/', skipDuplicates=true)")]
    public string Execute( // NOSONAR S107 - MCP protocol requires multiple parameters
        [Description("Operation: add, edit, delete, get, export_slides, extract")]
        string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Slide index (0-based, required for add/edit/delete/get)")]
        int? slideIndex = null,
        [Description("Image index on the slide (0-based, required for edit/delete)")]
        int? imageIndex = null,
        [Description("Image file path (required for add, optional for edit)")]
        string? imagePath = null,
        [Description("X position in points (optional for add/edit, default: 100)")]
        float x = 100,
        [Description("Y position in points (optional for add/edit, default: 100)")]
        float y = 100,
        [Description("Width in points (optional for add/edit)")]
        float? width = null,
        [Description("Height in points (optional for add/edit)")]
        float? height = null,
        [Description("JPEG quality 10-100 (optional for edit, re-encode image as JPEG)")]
        int? jpegQuality = null,
        [Description("Maximum width in pixels for resize (optional for edit)")]
        int? maxWidth = null,
        [Description("Maximum height in pixels for resize (optional for edit)")]
        int? maxHeight = null,
        [Description("Output directory (required for export_slides/extract)")]
        string? outputDir = null,
        [Description("Image format: png|jpeg (optional for export_slides/extract, default: png)")]
        string format = "png",
        [Description("Scaling factor (optional for export_slides, default: 1.0)")]
        float scale = 1.0f,
        [Description("Comma-separated slide indexes to export (optional for export_slides, e.g., '0,2,4')")]
        string? slideIndexes = null,
        [Description("Skip duplicate images based on content hash (optional for extract, default: false)")]
        bool skipDuplicates = false)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, slideIndex, imageIndex, imagePath, x, y, width, height,
            jpegQuality, maxWidth, maxHeight, outputDir, format, scale, slideIndexes, skipDuplicates);

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

        var lowerOp = operation.ToLowerInvariant();
        if (lowerOp == "get" || lowerOp == "export_slides" || lowerOp == "extract")
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
    private static OperationParameters BuildParameters( // NOSONAR S107 - MCP protocol parameter building
        string operation,
        int? slideIndex,
        int? imageIndex,
        string? imagePath,
        float x,
        float y,
        float? width,
        float? height,
        int? jpegQuality,
        int? maxWidth,
        int? maxHeight,
        string? outputDir,
        string format,
        float scale,
        string? slideIndexes,
        bool skipDuplicates)
    {
        return operation.ToLowerInvariant() switch
        {
            "add" => BuildAddParameters(slideIndex, imagePath, x, y, width, height),
            "edit" => BuildEditParameters(slideIndex, imageIndex, imagePath, x, y, width, height, jpegQuality, maxWidth,
                maxHeight),
            "delete" => BuildDeleteParameters(slideIndex, imageIndex),
            "get" => BuildGetParameters(slideIndex),
            "export_slides" => BuildExportSlidesParameters(outputDir, slideIndexes, format, scale),
            "extract" => BuildExtractParameters(outputDir, format, skipDuplicates),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the add image operation.
    /// </summary>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="imagePath">The image file path to add.</param>
    /// <param name="x">The X position in points.</param>
    /// <param name="y">The Y position in points.</param>
    /// <param name="width">The width in points.</param>
    /// <param name="height">The height in points.</param>
    /// <returns>OperationParameters configured for adding an image.</returns>
    private static OperationParameters BuildAddParameters(int? slideIndex, string? imagePath, float x, float y,
        float? width, float? height)
    {
        var parameters = new OperationParameters();
        if (slideIndex.HasValue) parameters.Set("slideIndex", slideIndex.Value);
        if (imagePath != null) parameters.Set("imagePath", imagePath);
        parameters.Set("x", x);
        parameters.Set("y", y);
        if (width.HasValue) parameters.Set("width", width.Value);
        if (height.HasValue) parameters.Set("height", height.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the edit image operation.
    /// </summary>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="imageIndex">The image index on the slide (0-based).</param>
    /// <param name="imagePath">The new image file path.</param>
    /// <param name="x">The X position in points.</param>
    /// <param name="y">The Y position in points.</param>
    /// <param name="width">The width in points.</param>
    /// <param name="height">The height in points.</param>
    /// <param name="jpegQuality">The JPEG quality (10-100) for re-encoding.</param>
    /// <param name="maxWidth">The maximum width in pixels for resize.</param>
    /// <param name="maxHeight">The maximum height in pixels for resize.</param>
    /// <returns>OperationParameters configured for editing an image.</returns>
    private static OperationParameters BuildEditParameters(int? slideIndex, int? imageIndex, string? imagePath, float x,
        float y, float? width, float? height, int? jpegQuality, int? maxWidth, int? maxHeight)
    {
        var parameters = new OperationParameters();
        if (slideIndex.HasValue) parameters.Set("slideIndex", slideIndex.Value);
        if (imageIndex.HasValue) parameters.Set("imageIndex", imageIndex.Value);
        if (imagePath != null) parameters.Set("imagePath", imagePath);
        parameters.Set("x", x);
        parameters.Set("y", y);
        if (width.HasValue) parameters.Set("width", width.Value);
        if (height.HasValue) parameters.Set("height", height.Value);
        if (jpegQuality.HasValue) parameters.Set("jpegQuality", jpegQuality.Value);
        if (maxWidth.HasValue) parameters.Set("maxWidth", maxWidth.Value);
        if (maxHeight.HasValue) parameters.Set("maxHeight", maxHeight.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the delete image operation.
    /// </summary>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="imageIndex">The image index on the slide (0-based).</param>
    /// <returns>OperationParameters configured for deleting an image.</returns>
    private static OperationParameters BuildDeleteParameters(int? slideIndex, int? imageIndex)
    {
        var parameters = new OperationParameters();
        if (slideIndex.HasValue) parameters.Set("slideIndex", slideIndex.Value);
        if (imageIndex.HasValue) parameters.Set("imageIndex", imageIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the get image information operation.
    /// </summary>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <returns>OperationParameters configured for getting image information.</returns>
    private static OperationParameters BuildGetParameters(int? slideIndex)
    {
        var parameters = new OperationParameters();
        if (slideIndex.HasValue) parameters.Set("slideIndex", slideIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the export slides as images operation.
    /// </summary>
    /// <param name="outputDir">The output directory path.</param>
    /// <param name="slideIndexes">Comma-separated slide indexes to export.</param>
    /// <param name="format">The image format (png or jpeg).</param>
    /// <param name="scale">The scaling factor.</param>
    /// <returns>OperationParameters configured for exporting slides as images.</returns>
    private static OperationParameters BuildExportSlidesParameters(string? outputDir, string? slideIndexes,
        string format, float scale)
    {
        var parameters = new OperationParameters();
        if (outputDir != null) parameters.Set("outputDir", outputDir);
        if (slideIndexes != null) parameters.Set("slideIndexes", slideIndexes);
        parameters.Set("format", format);
        parameters.Set("scale", scale);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the extract embedded images operation.
    /// </summary>
    /// <param name="outputDir">The output directory path.</param>
    /// <param name="format">The image format (png or jpeg).</param>
    /// <param name="skipDuplicates">Whether to skip duplicate images based on content hash.</param>
    /// <returns>OperationParameters configured for extracting embedded images.</returns>
    private static OperationParameters BuildExtractParameters(string? outputDir, string format, bool skipDuplicates)
    {
        var parameters = new OperationParameters();
        if (outputDir != null) parameters.Set("outputDir", outputDir);
        parameters.Set("format", format);
        parameters.Set("skipDuplicates", skipDuplicates);
        return parameters;
    }
}
