using System.ComponentModel;
using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing headers and footers in PDF documents (add text, add image, add page numbers, remove)
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Pdf.HeaderFooter")]
[McpServerToolType]
public class PdfHeaderFooterTool
{
    /// <summary>
    ///     Handler registry for header/footer operations.
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
    ///     Initializes a new instance of the <see cref="PdfHeaderFooterTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public PdfHeaderFooterTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry =
            HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Pdf.HeaderFooter");
    }

    /// <summary>
    ///     Executes a PDF header/footer operation (add_text, add_image, add_page_number, remove).
    /// </summary>
    /// <param name="operation">The operation to perform: add_text, add_image, add_page_number, remove.</param>
    /// <param name="path">PDF file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="text">Text content for add_text operation.</param>
    /// <param name="imagePath">Image file path for add_image operation.</param>
    /// <param name="format">Page number format string ({0} = current, {1} = total).</param>
    /// <param name="position">Position: header or footer.</param>
    /// <param name="alignment">Horizontal alignment: left, center, right.</param>
    /// <param name="fontSize">Font size in points.</param>
    /// <param name="margin">Margin from edge in points.</param>
    /// <param name="width">Image width (0 = auto, for add_image).</param>
    /// <param name="height">Image height (0 = auto, for add_image).</param>
    /// <param name="startPage">Starting page number for page numbering.</param>
    /// <param name="pageRange">Page range (e.g., "1-3" or "1,3,5"). If not specified, applies to all pages.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "pdf_header_footer",
        Title = "PDF Header/Footer Operations",
        Destructive = false,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(
        @"Add text, images, or page numbers to PDF headers and footers. Supports 4 operations: add_text, add_image, add_page_number, remove.

Usage examples:
- Add text header: pdf_header_footer(operation='add_text', path='doc.pdf', outputPath='out.pdf', text='Confidential', position='header')
- Add footer image: pdf_header_footer(operation='add_image', path='doc.pdf', outputPath='out.pdf', imagePath='logo.png', position='footer')
- Add page numbers: pdf_header_footer(operation='add_page_number', path='doc.pdf', outputPath='out.pdf', position='footer', alignment='right')
- Remove stamps: pdf_header_footer(operation='remove', path='doc.pdf', outputPath='out.pdf')")]
    public object Execute(
        [Description("Operation: add_text, add_image, add_page_number, remove")]
        string operation,
        [Description("PDF file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Text content (required for add_text)")]
        string? text = null,
        [Description("Image file path (required for add_image)")]
        string? imagePath = null,
        [Description("Page number format string ({0} = current, {1} = total, default: 'Page {0} of {1}')")]
        string format = "Page {0} of {1}",
        [Description("Position: header or footer (default: header)")]
        string position = "header",
        [Description("Horizontal alignment: left, center, right (default: center)")]
        string alignment = "center",
        [Description("Font size in points (default: 12.0)")]
        double fontSize = 12.0,
        [Description("Margin from edge in points (default: 20.0)")]
        double margin = 20.0,
        [Description("Image width, 0 = auto (default: 0, for add_image)")]
        double width = 0,
        [Description("Image height, 0 = auto (default: 0, for add_image)")]
        double height = 0,
        [Description("Starting page number for page numbering (default: 1)")]
        int startPage = 1,
        [Description("Page range (e.g., '1-3' or '1,3,5'). If not specified, applies to all pages")]
        string? pageRange = null)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, text, imagePath, format, position, alignment,
            fontSize, margin, width, height, startPage, pageRange);

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

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return ResultHelper.FinalizeResult((dynamic)result, ctx, outputPath);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    /// </summary>
    /// <param name="operation">The operation being performed.</param>
    /// <param name="text">The text content.</param>
    /// <param name="imagePath">The image file path.</param>
    /// <param name="format">The page number format string.</param>
    /// <param name="position">The header/footer position.</param>
    /// <param name="alignment">The horizontal alignment.</param>
    /// <param name="fontSize">The font size in points.</param>
    /// <param name="margin">The margin from edge in points.</param>
    /// <param name="width">The image width.</param>
    /// <param name="height">The image height.</param>
    /// <param name="startPage">The starting page number.</param>
    /// <param name="pageRange">The page range string.</param>
    /// <returns>OperationParameters configured for the specified operation.</returns>
    private static OperationParameters BuildParameters(
        string operation,
        string? text,
        string? imagePath,
        string format,
        string position,
        string alignment,
        double fontSize,
        double margin,
        double width,
        double height,
        int startPage,
        string? pageRange)
    {
        return operation.ToLowerInvariant() switch
        {
            "add_text" => BuildAddTextParameters(text, position, alignment, fontSize, margin, pageRange),
            "add_image" => BuildAddImageParameters(imagePath, position, alignment, margin, width, height, pageRange),
            "add_page_number" => BuildAddPageNumberParameters(format, position, alignment, fontSize, margin, startPage,
                pageRange),
            "remove" => BuildRemoveParameters(pageRange),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the add text header/footer operation.
    /// </summary>
    /// <param name="text">The text content.</param>
    /// <param name="position">The header/footer position.</param>
    /// <param name="alignment">The horizontal alignment.</param>
    /// <param name="fontSize">The font size in points.</param>
    /// <param name="margin">The margin from edge in points.</param>
    /// <param name="pageRange">The page range string.</param>
    /// <returns>OperationParameters configured for adding text header/footer.</returns>
    private static OperationParameters BuildAddTextParameters(
        string? text,
        string position,
        string alignment,
        double fontSize,
        double margin,
        string? pageRange)
    {
        var parameters = new OperationParameters();
        if (text != null) parameters.Set("text", text);
        parameters.Set("position", position);
        parameters.Set("alignment", alignment);
        parameters.Set("fontSize", fontSize);
        parameters.Set("margin", margin);
        if (pageRange != null) parameters.Set("pageRange", pageRange);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the add image header/footer operation.
    /// </summary>
    /// <param name="imagePath">The image file path.</param>
    /// <param name="position">The header/footer position.</param>
    /// <param name="alignment">The horizontal alignment.</param>
    /// <param name="margin">The margin from edge in points.</param>
    /// <param name="width">The image width.</param>
    /// <param name="height">The image height.</param>
    /// <param name="pageRange">The page range string.</param>
    /// <returns>OperationParameters configured for adding image header/footer.</returns>
    private static OperationParameters BuildAddImageParameters(
        string? imagePath,
        string position,
        string alignment,
        double margin,
        double width,
        double height,
        string? pageRange)
    {
        var parameters = new OperationParameters();
        if (imagePath != null) parameters.Set("imagePath", imagePath);
        parameters.Set("position", position);
        parameters.Set("alignment", alignment);
        parameters.Set("margin", margin);
        parameters.Set("width", width);
        parameters.Set("height", height);
        if (pageRange != null) parameters.Set("pageRange", pageRange);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the add page number operation.
    /// </summary>
    /// <param name="format">The page number format string.</param>
    /// <param name="position">The header/footer position.</param>
    /// <param name="alignment">The horizontal alignment.</param>
    /// <param name="fontSize">The font size in points.</param>
    /// <param name="margin">The margin from edge in points.</param>
    /// <param name="startPage">The starting page number.</param>
    /// <param name="pageRange">The page range string.</param>
    /// <returns>OperationParameters configured for adding page numbers.</returns>
    private static OperationParameters BuildAddPageNumberParameters(
        string format,
        string position,
        string alignment,
        double fontSize,
        double margin,
        int startPage,
        string? pageRange)
    {
        var parameters = new OperationParameters();
        parameters.Set("format", format);
        parameters.Set("position", position);
        parameters.Set("alignment", alignment);
        parameters.Set("fontSize", fontSize);
        parameters.Set("margin", margin);
        parameters.Set("startPage", startPage);
        if (pageRange != null) parameters.Set("pageRange", pageRange);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the remove stamps operation.
    /// </summary>
    /// <param name="pageRange">The page range string.</param>
    /// <returns>OperationParameters configured for removing stamps.</returns>
    private static OperationParameters BuildRemoveParameters(string? pageRange)
    {
        var parameters = new OperationParameters();
        if (pageRange != null) parameters.Set("pageRange", pageRange);
        return parameters;
    }
}
