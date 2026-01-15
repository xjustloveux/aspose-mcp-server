using System.ComponentModel;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for managing Word document properties (get, set)
///     Merges: WordGetDocumentPropertiesTool, WordSetDocumentPropertiesTool, WordSetPropertiesTool
/// </summary>
[McpServerToolType]
public class WordPropertiesTool
{
    /// <summary>
    ///     Handler registry for properties operations
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
    ///     Initializes a new instance of the WordPropertiesTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordPropertiesTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Word.Properties");
    }

    /// <summary>
    ///     Executes a Word properties operation (get, set).
    /// </summary>
    /// <param name="operation">The operation to perform: get, set.</param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (for set operation).</param>
    /// <param name="title">Document title (for set).</param>
    /// <param name="subject">Document subject (for set).</param>
    /// <param name="author">Document author (for set).</param>
    /// <param name="keywords">Keywords (for set).</param>
    /// <param name="comments">Comments (for set).</param>
    /// <param name="category">Document category (for set).</param>
    /// <param name="company">Company name (for set).</param>
    /// <param name="manager">Manager name (for set).</param>
    /// <param name="customProperties">Custom properties as JSON string (for set).</param>
    /// <returns>Document properties as JSON for get, or a success message for set operations.</returns>
    /// <exception cref="ArgumentException">Thrown when the operation is unknown.</exception>
    [McpServerTool(Name = "word_properties")]
    [Description(@"Get or set Word document properties (metadata). Supports 2 operations: get, set.

Usage examples:
- Get properties: word_properties(operation='get', path='doc.docx')
- Set properties: word_properties(operation='set', path='doc.docx', title='Title', author='Author', subject='Subject')

Notes:
- The 'set' operation is for content metadata (title, author, subject, etc.), not for statistics (word count, page count)
- Statistics like word count and page count are automatically calculated by Word and cannot be manually set
- Custom properties support multiple types: string, number (integer/double), boolean, and datetime (ISO 8601 format)")]
    public string Execute(
        [Description("Operation: get, set")] string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (if not provided, overwrites input, for set operation)")]
        string? outputPath = null,
        [Description("Document title (optional, for set operation)")]
        string? title = null,
        [Description("Document subject (optional, for set operation)")]
        string? subject = null,
        [Description("Document author (optional, for set operation)")]
        string? author = null,
        [Description("Keywords (optional, for set operation)")]
        string? keywords = null,
        [Description("Comments (optional, for set operation)")]
        string? comments = null,
        [Description("Document category (optional, for set operation)")]
        string? category = null,
        [Description("Company name (optional, for set operation)")]
        string? company = null,
        [Description("Manager name (optional, for set operation)")]
        string? manager = null,
        [Description(
            "Custom properties as JSON string (optional, for set operation). Supports string, number (integer/double), boolean, and datetime (ISO 8601 format).")]
        string? customProperties = null)
    {
        var parameters = BuildParameters(operation, title, subject, author, keywords, comments, category, company,
            manager, customProperties);

        var handler = _handlerRegistry.GetHandler(operation);

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var effectiveOutputPath = outputPath ?? path;

        var operationContext = new OperationContext<Document>
        {
            Document = ctx.Document,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = path,
            OutputPath = effectiveOutputPath
        };

        var result = handler.Execute(operationContext, parameters);

        if (operationContext.IsModified)
            ctx.Save(effectiveOutputPath);

        return ctx.IsSession ? result :
            operationContext.IsModified ? $"{result}\n{ctx.GetOutputMessage(effectiveOutputPath)}" : result;
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters.
    /// </summary>
    private static OperationParameters BuildParameters(
        string operation,
        string? title,
        string? subject,
        string? author,
        string? keywords,
        string? comments,
        string? category,
        string? company,
        string? manager,
        string? customProperties)
    {
        var parameters = new OperationParameters();

        if (operation.ToLower() == "set")
        {
            if (title != null) parameters.Set("title", title);
            if (subject != null) parameters.Set("subject", subject);
            if (author != null) parameters.Set("author", author);
            if (keywords != null) parameters.Set("keywords", keywords);
            if (comments != null) parameters.Set("comments", comments);
            if (category != null) parameters.Set("category", category);
            if (company != null) parameters.Set("company", company);
            if (manager != null) parameters.Set("manager", manager);
            if (customProperties != null) parameters.Set("customProperties", customProperties);
        }

        return parameters;
    }
}
