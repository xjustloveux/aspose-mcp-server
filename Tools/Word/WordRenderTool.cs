using System.ComponentModel;
using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Tool for rendering Word document pages to images (render_page, render_thumbnail).
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Word.Render")]
[McpServerToolType]
public class WordRenderTool
{
    /// <summary>
    ///     Handler registry for render operations.
    /// </summary>
    private readonly HandlerRegistry<Document> _handlerRegistry;

    /// <summary>
    ///     The session identity accessor for session isolation.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     The document session manager for managing in-memory document sessions.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="WordRenderTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public WordRenderTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Word.Render");
    }

    /// <summary>
    ///     Executes a Word document render operation (render_page, render_thumbnail).
    /// </summary>
    /// <param name="operation">The operation to perform: render_page, render_thumbnail.</param>
    /// <param name="path">Word document file path (required).</param>
    /// <param name="outputPath">Output image file path (required).</param>
    /// <param name="pageIndex">Page index (1-based, for render_page; default: all pages).</param>
    /// <param name="format">Image format: png, jpeg, bmp, tiff, svg (default: png).</param>
    /// <param name="dpi">Rendering DPI (for render_page, default: 150).</param>
    /// <param name="scale">Scale factor for thumbnail (0-1, for render_thumbnail, default: 0.25).</param>
    /// <returns>Render result with output file paths.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "word_render",
        Title = "Word Render Operations",
        Destructive = false,
        Idempotent = true,
        OpenWorld = false,
        ReadOnly = true,
        UseStructuredContent = true)]
    [Description(@"Render Word document pages to images. Supports 2 operations: render_page, render_thumbnail.

Usage examples:
- Render single page: word_render(operation='render_page', path='doc.docx', outputPath='page.png', pageIndex=1)
- Render all pages: word_render(operation='render_page', path='doc.docx', outputPath='output/pages.png')
- Render thumbnail: word_render(operation='render_thumbnail', path='doc.docx', outputPath='thumb.png')
- Render with options: word_render(operation='render_page', path='doc.docx', outputPath='page.jpeg', format='jpeg', dpi=300)")]
    public object Execute(
        [Description(@"Operation to perform.
- 'render_page': Render page(s) to image (required params: path, outputPath)
- 'render_thumbnail': Render first page thumbnail (required params: path, outputPath)")]
        string operation,
        [Description("Word document file path (required)")]
        string? path = null,
        [Description("Output image file path (required)")]
        string? outputPath = null,
        [Description("Page index (1-based, for render_page; omit to render all pages)")]
        int? pageIndex = null,
        [Description("Image format: png, jpeg, bmp, tiff, svg (default: png)")]
        string format = "png",
        [Description("Rendering DPI (for render_page, default: 150)")]
        int dpi = 150,
        [Description("Scale factor for thumbnail (0-1, for render_thumbnail, default: 0.25)")]
        double scale = 0.25)
    {
        var parameters = BuildParameters(operation, path, outputPath, pageIndex, format, dpi, scale);

        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<Document>
        {
            Document = null!,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SourcePath = path,
            OutputPath = outputPath
        };

        var result = handler.Execute(operationContext, parameters);
        return ResultHelper.FinalizeResult((dynamic)result, outputPath, (string?)null);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters(
        string operation,
        string? path,
        string? outputPath,
        int? pageIndex,
        string format,
        int dpi,
        double scale)
    {
        return operation.ToLowerInvariant() switch
        {
            "render_page" => BuildRenderPageParameters(path, outputPath, pageIndex, format, dpi),
            "render_thumbnail" => BuildRenderThumbnailParameters(path, outputPath, format, scale),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the render_page operation.
    /// </summary>
    /// <param name="path">The source document file path.</param>
    /// <param name="outputPath">The output image file path.</param>
    /// <param name="pageIndex">The 1-based page index.</param>
    /// <param name="format">The output image format.</param>
    /// <param name="dpi">The rendering DPI.</param>
    /// <returns>OperationParameters configured for rendering pages.</returns>
    private static OperationParameters BuildRenderPageParameters(string? path, string? outputPath,
        int? pageIndex, string format, int dpi)
    {
        var parameters = new OperationParameters();
        if (path != null) parameters.Set("path", path);
        if (outputPath != null) parameters.Set("outputPath", outputPath);
        if (pageIndex.HasValue) parameters.Set("pageIndex", pageIndex.Value);
        parameters.Set("format", format);
        parameters.Set("dpi", dpi);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the render_thumbnail operation.
    /// </summary>
    /// <param name="path">The source document file path.</param>
    /// <param name="outputPath">The output image file path.</param>
    /// <param name="format">The output image format.</param>
    /// <param name="scale">The scale factor.</param>
    /// <returns>OperationParameters configured for rendering thumbnails.</returns>
    private static OperationParameters BuildRenderThumbnailParameters(string? path, string? outputPath,
        string format, double scale)
    {
        var parameters = new OperationParameters();
        if (path != null) parameters.Set("path", path);
        if (outputPath != null) parameters.Set("outputPath", outputPath);
        parameters.Set("format", format);
        parameters.Set("scale", scale);
        return parameters;
    }
}
