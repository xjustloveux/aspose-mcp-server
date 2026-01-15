using System.ComponentModel;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Tool for getting Word document content, statistics, and document info
/// </summary>
[McpServerToolType]
public class WordContentTool
{
    /// <summary>
    ///     Handler registry for content operations
    /// </summary>
    private readonly HandlerRegistry<Document> _handlerRegistry;

    /// <summary>
    ///     Identity accessor for session isolation
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document session operations
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the WordContentTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordContentTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Word.Content");
    }

    /// <summary>
    ///     Executes a Word content operation (get_content, get_content_detailed, get_statistics, get_document_info).
    /// </summary>
    /// <param name="operation">The operation to perform: get_content, get_content_detailed, get_statistics, get_document_info.</param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="includeHeaders">Include headers in content (for get_content_detailed, default: false).</param>
    /// <param name="includeFooters">Include footers in content (for get_content_detailed, default: false).</param>
    /// <param name="includeFootnotes">Include footnotes in statistics (for get_statistics, default: true).</param>
    /// <param name="includeTabStops">Include tab stops in document info (for get_document_info, default: false).</param>
    /// <param name="maxChars">Maximum characters to return (for get_content/get_content_detailed).</param>
    /// <param name="offset">Character offset to start reading from (for get_content/get_content_detailed, default: 0).</param>
    /// <returns>Document content, detailed content, statistics, or document info as string or JSON.</returns>
    /// <exception cref="ArgumentException">Thrown when the operation is unknown.</exception>
    [McpServerTool(Name = "word_content")]
    [Description(
        @"Get Word document content, statistics, and document information. Supports 4 operations: get_content, get_content_detailed, get_statistics, get_document_info.

Usage examples:
- Get content: word_content(operation='get_content', path='doc.docx')
- Get detailed content: word_content(operation='get_content_detailed', path='doc.docx', includeHeaders=true, includeFooters=true)
- Get statistics: word_content(operation='get_statistics', path='doc.docx', includeFootnotes=true)
- Get document info: word_content(operation='get_document_info', path='doc.docx', includeTabStops=true)")]
    public string Execute(
        [Description("Operation: get_content, get_content_detailed, get_statistics, get_document_info")]
        string operation,
        [Description("Word document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Include headers in content (for get_content_detailed, default: false)")]
        bool includeHeaders = false,
        [Description("Include footers in content (for get_content_detailed, default: false)")]
        bool includeFooters = false,
        [Description("Include footnotes in statistics (for get_statistics, default: true)")]
        bool includeFootnotes = true,
        [Description("Include tab stops in document info (for get_document_info, default: false)")]
        bool includeTabStops = false,
        [Description(
            "Maximum characters to return (for get_content/get_content_detailed). Use for large documents to avoid token overflow.")]
        int? maxChars = null,
        [Description("Character offset to start reading from (for get_content/get_content_detailed, default: 0)")]
        int offset = 0)
    {
        var parameters = BuildParameters(operation, includeHeaders, includeFooters, includeFootnotes, includeTabStops,
            maxChars, offset);

        var handler = _handlerRegistry.GetHandler(operation);

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var operationContext = new OperationContext<Document>
        {
            Document = ctx.Document,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = path,
            OutputPath = path
        };

        return handler.Execute(operationContext, parameters);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters.
    /// </summary>
    private static OperationParameters BuildParameters(
        string operation,
        bool includeHeaders,
        bool includeFooters,
        bool includeFootnotes,
        bool includeTabStops,
        int? maxChars,
        int offset)
    {
        var parameters = new OperationParameters();

        switch (operation.ToLower())
        {
            case "get_content":
                if (maxChars.HasValue) parameters.Set("maxChars", maxChars.Value);
                parameters.Set("offset", offset);
                break;

            case "get_content_detailed":
                parameters.Set("includeHeaders", includeHeaders);
                parameters.Set("includeFooters", includeFooters);
                break;

            case "get_statistics":
                parameters.Set("includeFootnotes", includeFootnotes);
                break;

            case "get_document_info":
                parameters.Set("includeTabStops", includeTabStops);
                break;
        }

        return parameters;
    }
}
