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
    ///     Builds OperationParameters from method parameters.
    /// </summary>
    private static OperationParameters BuildParameters(
        string operation,
        int? pageIndex,
        int? annotationIndex,
        string? text,
        double x,
        double y)
    {
        var parameters = new OperationParameters();

        switch (operation.ToLowerInvariant())
        {
            case "add":
                if (pageIndex.HasValue) parameters.Set("pageIndex", pageIndex.Value);
                if (text != null) parameters.Set("text", text);
                parameters.Set("x", x);
                parameters.Set("y", y);
                break;

            case "delete":
                if (pageIndex.HasValue) parameters.Set("pageIndex", pageIndex.Value);
                if (annotationIndex.HasValue) parameters.Set("annotationIndex", annotationIndex.Value);
                break;

            case "edit":
                if (pageIndex.HasValue) parameters.Set("pageIndex", pageIndex.Value);
                if (annotationIndex.HasValue) parameters.Set("annotationIndex", annotationIndex.Value);
                if (text != null) parameters.Set("text", text);
                parameters.Set("x", x);
                parameters.Set("y", y);
                break;

            case "get":
                if (pageIndex.HasValue) parameters.Set("pageIndex", pageIndex.Value);
                break;
        }

        return parameters;
    }
}
