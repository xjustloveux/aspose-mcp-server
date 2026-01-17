using System.ComponentModel;
using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing annotations in PDF documents (add, delete, edit, get)
/// </summary>
[McpServerToolType]
public class PdfAnnotationTool
{
    /// <summary>
    ///     Handler registry for annotation operations.
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
    ///     Initializes a new instance of the <see cref="PdfAnnotationTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public PdfAnnotationTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Pdf.Annotation");
    }

    /// <summary>
    ///     Executes a PDF annotation operation (add, delete, edit, get).
    /// </summary>
    /// <param name="operation">The operation to perform: add, delete, edit, get.</param>
    /// <param name="path">PDF file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="pageIndex">Page index (1-based, required for add, delete, edit).</param>
    /// <param name="annotationIndex">Annotation index (1-based, required for edit, optional for delete).</param>
    /// <param name="text">Annotation text (required for add, edit).</param>
    /// <param name="x">X position in points (origin is bottom-left, 72 points = 1 inch).</param>
    /// <param name="y">Y position in points (origin is bottom-left, 72 points = 1 inch).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "pdf_annotation")]
    [Description(@"Manage annotations in PDF documents. Supports 4 operations: add, delete, edit, get.

Usage examples:
- Add annotation: pdf_annotation(operation='add', path='doc.pdf', pageIndex=1, text='Note', x=100, y=100)
- Delete annotation: pdf_annotation(operation='delete', path='doc.pdf', pageIndex=1, annotationIndex=1)
- Edit annotation: pdf_annotation(operation='edit', path='doc.pdf', pageIndex=1, annotationIndex=1, text='Updated Note')
- Get annotations: pdf_annotation(operation='get', path='doc.pdf', pageIndex=1)")]
    public string Execute(
        [Description(@"Operation to perform.
- 'add': Add an annotation (required params: path, pageIndex, text, x, y)
- 'delete': Delete annotation(s) (required params: path, pageIndex; optional: annotationIndex, deletes all if omitted)
- 'edit': Edit an annotation (required params: path, pageIndex, annotationIndex, text)
- 'get': Get all annotations (required params: path, pageIndex)")]
        string operation,
        [Description("PDF file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Page index (1-based, required for add, delete, edit)")]
        int? pageIndex = null,
        [Description("Annotation index (1-based, required for edit, optional for delete - deletes all if omitted)")]
        int? annotationIndex = null,
        [Description("Annotation text (required for add, edit)")]
        string? text = null,
        [Description("X position in points (origin is bottom-left, 72 points = 1 inch, for add/edit, default: 100)")]
        double x = 100,
        [Description("Y position in points (origin is bottom-left, 72 points = 1 inch, for add/edit, default: 700)")]
        double y = 700)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, pageIndex, annotationIndex, text, x, y);

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
        int? pageIndex,
        int? annotationIndex,
        string? text,
        double x,
        double y)
    {
        return operation.ToLowerInvariant() switch
        {
            "add" => BuildAddParameters(pageIndex, text, x, y),
            "delete" => BuildDeleteParameters(pageIndex, annotationIndex),
            "edit" => BuildEditParameters(pageIndex, annotationIndex, text, x, y),
            "get" => BuildGetParameters(pageIndex),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the add annotation operation.
    /// </summary>
    /// <param name="pageIndex">The page index (1-based) to add the annotation to.</param>
    /// <param name="text">The annotation text content.</param>
    /// <param name="x">The X position in points.</param>
    /// <param name="y">The Y position in points.</param>
    /// <returns>OperationParameters configured for adding an annotation.</returns>
    private static OperationParameters BuildAddParameters(int? pageIndex, string? text, double x, double y)
    {
        var parameters = new OperationParameters();
        if (pageIndex.HasValue) parameters.Set("pageIndex", pageIndex.Value);
        if (text != null) parameters.Set("text", text);
        parameters.Set("x", x);
        parameters.Set("y", y);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the delete annotation operation.
    /// </summary>
    /// <param name="pageIndex">The page index (1-based) containing the annotation.</param>
    /// <param name="annotationIndex">The annotation index (1-based) to delete.</param>
    /// <returns>OperationParameters configured for deleting an annotation.</returns>
    private static OperationParameters BuildDeleteParameters(int? pageIndex, int? annotationIndex)
    {
        var parameters = new OperationParameters();
        if (pageIndex.HasValue) parameters.Set("pageIndex", pageIndex.Value);
        if (annotationIndex.HasValue) parameters.Set("annotationIndex", annotationIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the edit annotation operation.
    /// </summary>
    /// <param name="pageIndex">The page index (1-based) containing the annotation.</param>
    /// <param name="annotationIndex">The annotation index (1-based) to edit.</param>
    /// <param name="text">The new annotation text content.</param>
    /// <param name="x">The new X position in points.</param>
    /// <param name="y">The new Y position in points.</param>
    /// <returns>OperationParameters configured for editing an annotation.</returns>
    private static OperationParameters BuildEditParameters(int? pageIndex, int? annotationIndex, string? text,
        double x, double y)
    {
        var parameters = new OperationParameters();
        if (pageIndex.HasValue) parameters.Set("pageIndex", pageIndex.Value);
        if (annotationIndex.HasValue) parameters.Set("annotationIndex", annotationIndex.Value);
        if (text != null) parameters.Set("text", text);
        parameters.Set("x", x);
        parameters.Set("y", y);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the get annotations operation.
    /// </summary>
    /// <param name="pageIndex">The page index (1-based) to get annotations from.</param>
    /// <returns>OperationParameters configured for getting annotations.</returns>
    private static OperationParameters BuildGetParameters(int? pageIndex)
    {
        var parameters = new OperationParameters();
        if (pageIndex.HasValue) parameters.Set("pageIndex", pageIndex.Value);
        return parameters;
    }
}
