using System.ComponentModel;
using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing stamps in PDF documents (add_text, add_image, add_pdf, list, remove)
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Pdf.Stamp")]
[McpServerToolType]
public class PdfStampTool
{
    /// <summary>
    ///     Handler registry for stamp operations.
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
    ///     Initializes a new instance of the <see cref="PdfStampTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public PdfStampTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Pdf.Stamp");
    }

    /// <summary>
    ///     Executes a PDF stamp operation (add_text, add_image, add_pdf, list, remove).
    /// </summary>
    /// <param name="operation">The operation to perform: add_text, add_image, add_pdf, list, remove.</param>
    /// <param name="path">PDF file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="text">Stamp text content (required for add_text).</param>
    /// <param name="imagePath">Image file path (required for add_image).</param>
    /// <param name="pdfPath">Source PDF file path (required for add_pdf).</param>
    /// <param name="stampPageIndex">Page index from source PDF to use as stamp (1-based, for add_pdf).</param>
    /// <param name="pageIndex">Target page index (1-based, 0 = all pages for add operations, required for remove).</param>
    /// <param name="x">X position indent in points.</param>
    /// <param name="y">Y position indent in points.</param>
    /// <param name="width">Stamp width in points (0 = auto, for add_image, add_pdf).</param>
    /// <param name="height">Stamp height in points (0 = auto, for add_image, add_pdf).</param>
    /// <param name="fontSize">Font size for text stamp (for add_text).</param>
    /// <param name="opacity">Stamp opacity from 0.0 (transparent) to 1.0 (opaque).</param>
    /// <param name="rotation">Rotation angle in degrees.</param>
    /// <param name="color">Text color name or hex value (for add_text).</param>
    /// <param name="stampIndex">Stamp annotation index (1-based, for remove; omit to remove all stamps on page).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for list operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "pdf_stamp",
        Title = "PDF Stamp Operations",
        Destructive = true,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(
        @"Add text, image, or PDF page stamps to PDF documents. Supports 5 operations: add_text, add_image, add_pdf, list, remove.

Usage examples:
- Add text stamp: pdf_stamp(operation='add_text', path='doc.pdf', outputPath='out.pdf', text='CONFIDENTIAL', opacity=0.5, rotation=45)
- Add image stamp: pdf_stamp(operation='add_image', path='doc.pdf', outputPath='out.pdf', imagePath='logo.png', x=100, y=100)
- Add PDF stamp: pdf_stamp(operation='add_pdf', path='doc.pdf', outputPath='out.pdf', pdfPath='stamp.pdf', stampPageIndex=1)
- List stamps: pdf_stamp(operation='list', path='doc.pdf')
- Remove stamp: pdf_stamp(operation='remove', path='doc.pdf', outputPath='out.pdf', pageIndex=1, stampIndex=1)")]
    public object Execute(
        [Description(@"Operation to perform.
- 'add_text': Add a text stamp (required params: path, text)
- 'add_image': Add an image stamp (required params: path, imagePath)
- 'add_pdf': Add a PDF page stamp (required params: path, pdfPath)
- 'list': List stamp annotations (required params: path)
- 'remove': Remove stamp annotation(s) (required params: path, pageIndex; optional: stampIndex, removes all if omitted)")]
        string operation,
        [Description("PDF file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Stamp text content (required for add_text)")]
        string? text = null,
        [Description("Image file path (required for add_image)")]
        string? imagePath = null,
        [Description("Source PDF file path to use as stamp (required for add_pdf)")]
        string? pdfPath = null,
        [Description("Page index from source PDF to use as stamp (1-based, default: 1, for add_pdf)")]
        int stampPageIndex = 1,
        [Description("Target page index (1-based, 0 = all pages for add operations, required for remove)")]
        int pageIndex = 0,
        [Description("X position indent in points (default: 0)")]
        double x = 0,
        [Description("Y position indent in points (default: 0)")]
        double y = 0,
        [Description("Stamp width in points (0 = auto, for add_image, add_pdf)")]
        double width = 0,
        [Description("Stamp height in points (0 = auto, for add_image, add_pdf)")]
        double height = 0,
        [Description("Font size for text stamp (default: 14.0, for add_text)")]
        double fontSize = 14.0,
        [Description("Stamp opacity from 0.0 (transparent) to 1.0 (opaque, default)")]
        double opacity = 1.0,
        [Description("Rotation angle in degrees (default: 0)")]
        double rotation = 0.0,
        [Description("Text color name or hex value (default: 'black', for add_text)")]
        string color = "black",
        [Description("Stamp annotation index (1-based, for remove; omit to remove all stamps on page)")]
        int? stampIndex = null)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, text, imagePath, pdfPath, stampPageIndex, pageIndex,
            x, y, width, height, fontSize, opacity, rotation, color, stampIndex);

        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<Document>
        {
            Document = ctx.Document,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = path,
            OutputPath = outputPath
        };

        var result = handler.Execute(operationContext, parameters);

        if (string.Equals(operation, "list", StringComparison.OrdinalIgnoreCase))
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
        string? text,
        string? imagePath,
        string? pdfPath,
        int stampPageIndex,
        int pageIndex,
        double x,
        double y,
        double width,
        double height,
        double fontSize,
        double opacity,
        double rotation,
        string color,
        int? stampIndex)
    {
        return operation.ToLowerInvariant() switch
        {
            "add_text" => BuildAddTextParameters(text, pageIndex, x, y, fontSize, opacity, rotation, color),
            "add_image" => BuildAddImageParameters(imagePath, pageIndex, x, y, width, height, opacity, rotation),
            "add_pdf" => BuildAddPdfParameters(pdfPath, stampPageIndex, pageIndex, x, y, width, height, opacity,
                rotation),
            "list" => BuildListParameters(pageIndex),
            "remove" => BuildRemoveParameters(pageIndex, stampIndex),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the add text stamp operation.
    /// </summary>
    /// <param name="text">The stamp text content.</param>
    /// <param name="pageIndex">The target page index (0 = all pages).</param>
    /// <param name="x">The X position indent.</param>
    /// <param name="y">The Y position indent.</param>
    /// <param name="fontSize">The font size.</param>
    /// <param name="opacity">The stamp opacity.</param>
    /// <param name="rotation">The rotation angle in degrees.</param>
    /// <param name="color">The text color.</param>
    /// <returns>OperationParameters configured for adding a text stamp.</returns>
    private static OperationParameters BuildAddTextParameters(string? text, int pageIndex, double x, double y,
        double fontSize, double opacity, double rotation, string color)
    {
        var parameters = new OperationParameters();
        if (text != null) parameters.Set("text", text);
        parameters.Set("pageIndex", pageIndex);
        parameters.Set("x", x);
        parameters.Set("y", y);
        parameters.Set("fontSize", fontSize);
        parameters.Set("opacity", opacity);
        parameters.Set("rotation", rotation);
        parameters.Set("color", color);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the add image stamp operation.
    /// </summary>
    /// <param name="imagePath">The image file path.</param>
    /// <param name="pageIndex">The target page index (0 = all pages).</param>
    /// <param name="x">The X position indent.</param>
    /// <param name="y">The Y position indent.</param>
    /// <param name="width">The stamp width.</param>
    /// <param name="height">The stamp height.</param>
    /// <param name="opacity">The stamp opacity.</param>
    /// <param name="rotation">The rotation angle in degrees.</param>
    /// <returns>OperationParameters configured for adding an image stamp.</returns>
    private static OperationParameters BuildAddImageParameters(string? imagePath, int pageIndex, double x, double y,
        double width, double height, double opacity, double rotation)
    {
        var parameters = new OperationParameters();
        if (imagePath != null) parameters.Set("imagePath", imagePath);
        parameters.Set("pageIndex", pageIndex);
        parameters.Set("x", x);
        parameters.Set("y", y);
        parameters.Set("width", width);
        parameters.Set("height", height);
        parameters.Set("opacity", opacity);
        parameters.Set("rotation", rotation);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the add PDF page stamp operation.
    /// </summary>
    /// <param name="pdfPath">The source PDF file path.</param>
    /// <param name="stampPageIndex">The page index from the source PDF.</param>
    /// <param name="pageIndex">The target page index (0 = all pages).</param>
    /// <param name="x">The X position indent.</param>
    /// <param name="y">The Y position indent.</param>
    /// <param name="width">The stamp width.</param>
    /// <param name="height">The stamp height.</param>
    /// <param name="opacity">The stamp opacity.</param>
    /// <param name="rotation">The rotation angle in degrees.</param>
    /// <returns>OperationParameters configured for adding a PDF page stamp.</returns>
    private static OperationParameters BuildAddPdfParameters(string? pdfPath, int stampPageIndex, int pageIndex,
        double x, double y, double width, double height, double opacity, double rotation)
    {
        var parameters = new OperationParameters();
        if (pdfPath != null) parameters.Set("pdfPath", pdfPath);
        parameters.Set("stampPageIndex", stampPageIndex);
        parameters.Set("pageIndex", pageIndex);
        parameters.Set("x", x);
        parameters.Set("y", y);
        parameters.Set("width", width);
        parameters.Set("height", height);
        parameters.Set("opacity", opacity);
        parameters.Set("rotation", rotation);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the list stamps operation.
    /// </summary>
    /// <param name="pageIndex">The page index (0 = all pages).</param>
    /// <returns>OperationParameters configured for listing stamps.</returns>
    private static OperationParameters BuildListParameters(int pageIndex)
    {
        var parameters = new OperationParameters();
        parameters.Set("pageIndex", pageIndex);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the remove stamp operation.
    /// </summary>
    /// <param name="pageIndex">The page index (1-based).</param>
    /// <param name="stampIndex">The stamp annotation index (1-based, null = remove all).</param>
    /// <returns>OperationParameters configured for removing stamps.</returns>
    private static OperationParameters BuildRemoveParameters(int pageIndex, int? stampIndex)
    {
        var parameters = new OperationParameters();
        parameters.Set("pageIndex", pageIndex);
        if (stampIndex.HasValue) parameters.Set("stampIndex", stampIndex.Value);
        return parameters;
    }
}
