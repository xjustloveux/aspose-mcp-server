using System.ComponentModel;
using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing images in PDF documents (add, delete, edit, extract, get)
/// </summary>
[McpServerToolType]
public class PdfImageTool
{
    /// <summary>
    ///     Handler registry for image operations.
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
    ///     Initializes a new instance of the <see cref="PdfImageTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public PdfImageTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Pdf.Image");
    }

    /// <summary>
    ///     Executes a PDF image operation (add, delete, edit, extract, get).
    /// </summary>
    /// <param name="operation">The operation to perform: add, delete, edit, extract, get.</param>
    /// <param name="path">PDF file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (optional for add/delete/edit, required for extract).</param>
    /// <param name="pageIndex">Page index (1-based, required for add, delete, edit, extract, get).</param>
    /// <param name="imagePath">Image file path (required for add, optional for edit).</param>
    /// <param name="imageIndex">Image index (1-based, required for delete, edit, extract).</param>
    /// <param name="x">X position in PDF coordinates (for add, edit).</param>
    /// <param name="y">Y position in PDF coordinates (for add, edit).</param>
    /// <param name="width">Image width (for add, edit, optional).</param>
    /// <param name="height">Image height (for add, edit, optional).</param>
    /// <param name="outputDir">Output directory for extracted images (for extract).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "pdf_image")]
    [Description(@"Manage images in PDF documents. Supports 5 operations: add, delete, edit, extract, get.

Usage examples:
- Add image: pdf_image(operation='add', path='doc.pdf', pageIndex=1, imagePath='image.png', x=100, y=100)
- Delete image: pdf_image(operation='delete', path='doc.pdf', pageIndex=1, imageIndex=1)
- Move image: pdf_image(operation='edit', path='doc.pdf', pageIndex=1, imageIndex=1, x=200, y=200)
- Replace image: pdf_image(operation='edit', path='doc.pdf', pageIndex=1, imageIndex=1, imagePath='new.png', x=200, y=200)
- Extract image: pdf_image(operation='extract', path='doc.pdf', pageIndex=1, imageIndex=1, outputPath='image.png')
- Get images: pdf_image(operation='get', path='doc.pdf', pageIndex=1)")]
    public string Execute(
        [Description(@"Operation to perform.
- 'add': Add an image (required params: path, pageIndex, imagePath)
- 'delete': Delete an image (required params: path, pageIndex, imageIndex)
- 'edit': Edit image position/size (required params: path, pageIndex, imageIndex)
- 'extract': Extract an image (required params: path, pageIndex, imageIndex, outputPath)
- 'get': Get all images on a page (required params: path, pageIndex)")]
        string operation,
        [Description("PDF file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description(
            "Output file path (optional, defaults to overwrite input for add/delete/edit, required for extract)")]
        string? outputPath = null,
        [Description("Page index (1-based, required for add, delete, edit, extract, get)")]
        int pageIndex = 0,
        [Description("Image file path (required for add, optional for edit - omit to move existing image)")]
        string? imagePath = null,
        [Description("Image index (1-based, required for delete, edit, extract)")]
        int imageIndex = 0,
        [Description("X position in PDF coordinates, origin at bottom-left corner (for add, edit, default: 100)")]
        double x = 100,
        [Description("Y position in PDF coordinates, origin at bottom-left corner (for add, edit, default: 600)")]
        double y = 600,
        [Description("Image width (for add, edit, optional - if omitted defaults to 200)")]
        double? width = null,
        [Description("Image height (for add, edit, optional - if omitted defaults to 200)")]
        double? height = null,
        [Description("Output directory for extracted images (for extract)")]
        string? outputDir = null)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, pageIndex, imagePath, imageIndex, x, y, width, height, outputPath,
            outputDir);

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

        if (operation.Equals("get", StringComparison.OrdinalIgnoreCase) ||
            operation.Equals("extract", StringComparison.OrdinalIgnoreCase))
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
        int pageIndex,
        string? imagePath,
        int imageIndex,
        double x,
        double y,
        double? width,
        double? height,
        string? outputPath,
        string? outputDir)
    {
        return operation.ToLowerInvariant() switch
        {
            "add" => BuildAddParameters(pageIndex, imagePath, x, y, width, height),
            "delete" => BuildDeleteParameters(pageIndex, imageIndex),
            "edit" => BuildEditParameters(pageIndex, imageIndex, imagePath, x, y, width, height),
            "extract" => BuildExtractParameters(pageIndex, imageIndex, outputPath, outputDir),
            "get" => BuildGetParameters(pageIndex),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the add image operation.
    /// </summary>
    /// <param name="pageIndex">The page index (1-based) to add the image to.</param>
    /// <param name="imagePath">The file path of the image to add.</param>
    /// <param name="x">The X position in PDF coordinates.</param>
    /// <param name="y">The Y position in PDF coordinates.</param>
    /// <param name="width">The image width.</param>
    /// <param name="height">The image height.</param>
    /// <returns>OperationParameters configured for adding an image.</returns>
    private static OperationParameters BuildAddParameters(int pageIndex, string? imagePath, double x, double y,
        double? width, double? height)
    {
        var parameters = new OperationParameters();
        parameters.Set("pageIndex", pageIndex);
        if (imagePath != null) parameters.Set("imagePath", imagePath);
        parameters.Set("x", x);
        parameters.Set("y", y);
        if (width.HasValue) parameters.Set("width", width.Value);
        if (height.HasValue) parameters.Set("height", height.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the delete image operation.
    /// </summary>
    /// <param name="pageIndex">The page index (1-based) containing the image.</param>
    /// <param name="imageIndex">The image index (1-based) to delete.</param>
    /// <returns>OperationParameters configured for deleting an image.</returns>
    private static OperationParameters BuildDeleteParameters(int pageIndex, int imageIndex)
    {
        var parameters = new OperationParameters();
        parameters.Set("pageIndex", pageIndex);
        parameters.Set("imageIndex", imageIndex);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the edit image operation.
    /// </summary>
    /// <param name="pageIndex">The page index (1-based) containing the image.</param>
    /// <param name="imageIndex">The image index (1-based) to edit.</param>
    /// <param name="imagePath">The file path of the replacement image (optional).</param>
    /// <param name="x">The new X position in PDF coordinates.</param>
    /// <param name="y">The new Y position in PDF coordinates.</param>
    /// <param name="width">The new image width.</param>
    /// <param name="height">The new image height.</param>
    /// <returns>OperationParameters configured for editing an image.</returns>
    private static OperationParameters BuildEditParameters(int pageIndex, int imageIndex, string? imagePath, double x,
        double y, double? width, double? height)
    {
        var parameters = new OperationParameters();
        parameters.Set("pageIndex", pageIndex);
        parameters.Set("imageIndex", imageIndex);
        if (imagePath != null) parameters.Set("imagePath", imagePath);
        parameters.Set("x", x);
        parameters.Set("y", y);
        if (width.HasValue) parameters.Set("width", width.Value);
        if (height.HasValue) parameters.Set("height", height.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the extract image operation.
    /// </summary>
    /// <param name="pageIndex">The page index (1-based) containing the image.</param>
    /// <param name="imageIndex">The image index (1-based) to extract.</param>
    /// <param name="outputPath">The output file path for the extracted image.</param>
    /// <param name="outputDir">The output directory for extracted images.</param>
    /// <returns>OperationParameters configured for extracting an image.</returns>
    private static OperationParameters BuildExtractParameters(int pageIndex, int imageIndex, string? outputPath,
        string? outputDir)
    {
        var parameters = new OperationParameters();
        parameters.Set("pageIndex", pageIndex);
        if (imageIndex > 0) parameters.Set("imageIndex", imageIndex);
        if (outputPath != null) parameters.Set("outputPath", outputPath);
        if (outputDir != null) parameters.Set("outputDir", outputDir);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the get images operation.
    /// </summary>
    /// <param name="pageIndex">The page index (1-based) to get images from.</param>
    /// <returns>OperationParameters configured for getting images.</returns>
    private static OperationParameters BuildGetParameters(int pageIndex)
    {
        var parameters = new OperationParameters();
        if (pageIndex > 0) parameters.Set("pageIndex", pageIndex);
        return parameters;
    }
}
