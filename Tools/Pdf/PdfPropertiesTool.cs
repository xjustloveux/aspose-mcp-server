using System.ComponentModel;
using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing document properties in PDF files (get, set)
/// </summary>
[McpServerToolType]
public class PdfPropertiesTool
{
    /// <summary>
    ///     Handler registry for properties operations.
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
    ///     Initializes a new instance of the <see cref="PdfPropertiesTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public PdfPropertiesTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Pdf.Properties");
    }

    /// <summary>
    ///     Executes a PDF properties operation (get, set).
    /// </summary>
    /// <param name="operation">The operation to perform: get, set.</param>
    /// <param name="path">PDF file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (optional, defaults to overwrite input for set).</param>
    /// <param name="title">Title (for set).</param>
    /// <param name="author">Author (for set).</param>
    /// <param name="subject">Subject (for set).</param>
    /// <param name="keywords">Keywords (for set).</param>
    /// <param name="creator">Creator (for set).</param>
    /// <param name="producer">Producer (for set).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "pdf_properties")]
    [Description(@"Manage document properties in PDF files. Supports 2 operations: get, set.

Usage examples:
- Get properties: pdf_properties(operation='get', path='doc.pdf')
- Set properties: pdf_properties(operation='set', path='doc.pdf', title='Title', author='Author', subject='Subject')")]
    public string Execute(
        [Description(@"Operation to perform.
- 'get': Get document properties (required params: path)
- 'set': Set document properties (required params: path)")]
        string operation,
        [Description("PDF file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (optional, defaults to overwrite input for set)")]
        string? outputPath = null,
        [Description("Title (for set)")] string? title = null,
        [Description("Author (for set)")] string? author = null,
        [Description("Subject (for set)")] string? subject = null,
        [Description("Keywords (for set)")] string? keywords = null,
        [Description("Creator (for set)")] string? creator = null,
        [Description("Producer (for set)")] string? producer = null)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, title, author, subject, keywords, creator, producer);

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
        string? title,
        string? author,
        string? subject,
        string? keywords,
        string? creator,
        string? producer)
    {
        var parameters = new OperationParameters();

        if (string.Equals(operation, "set", StringComparison.OrdinalIgnoreCase))
        {
            if (title != null) parameters.Set("title", title);
            if (author != null) parameters.Set("author", author);
            if (subject != null) parameters.Set("subject", subject);
            if (keywords != null) parameters.Set("keywords", keywords);
            if (creator != null) parameters.Set("creator", creator);
            if (producer != null) parameters.Set("producer", producer);
        }

        return parameters;
    }
}
