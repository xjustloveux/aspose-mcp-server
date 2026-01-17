using System.ComponentModel;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for managing Word comments (add, delete, get, reply)
/// </summary>
[McpServerToolType]
public class WordCommentTool
{
    /// <summary>
    ///     Handler registry for comment operations
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
    ///     Initializes a new instance of the WordCommentTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordCommentTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Word.Comment");
    }

    /// <summary>
    ///     Executes a Word comment operation (add, delete, get, reply).
    /// </summary>
    /// <param name="operation">The operation to perform: add, delete, get, reply.</param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="text">Comment text content.</param>
    /// <param name="author">Comment author name.</param>
    /// <param name="authorInitial">Author initials.</param>
    /// <param name="paragraphIndex">Paragraph index (0-based).</param>
    /// <param name="startRunIndex">Start run index.</param>
    /// <param name="endRunIndex">End run index.</param>
    /// <param name="commentIndex">Comment index (0-based).</param>
    /// <param name="replyText">Reply text content.</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "word_comment")]
    [Description(@"Manage Word comments. Supports 4 operations: add, delete, get, reply.

Usage examples:
- Add comment: word_comment(operation='add', path='doc.docx', text='This is a comment', author='Author Name')
- Delete comment: word_comment(operation='delete', path='doc.docx', commentIndex=0)
- Get all comments: word_comment(operation='get', path='doc.docx')
- Reply to comment: word_comment(operation='reply', path='doc.docx', commentIndex=0, text='This is a reply')")]
    public string Execute( // NOSONAR S107 - MCP protocol requires multiple parameters
        [Description("Operation: add, delete, get, reply")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Comment text content")] string? text = null,
        [Description("Comment author name")] string? author = null,
        [Description("Author initials")] string? authorInitial = null,
        [Description("Paragraph index (0-based)")]
        int? paragraphIndex = null,
        [Description("Start run index")] int? startRunIndex = null,
        [Description("End run index")] int? endRunIndex = null,
        [Description("Comment index (0-based)")]
        int? commentIndex = null,
        [Description("Reply text content")] string? replyText = null)
    {
        var parameters = BuildParameters(operation, text, author, authorInitial, paragraphIndex, startRunIndex,
            endRunIndex, commentIndex, replyText);

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

        if (ctx.IsSession || !operationContext.IsModified)
            return result;
        return $"{result}\n{ctx.GetOutputMessage(effectiveOutputPath)}";
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters( // NOSONAR S107 - MCP protocol parameter building
        string operation,
        string? text,
        string? author,
        string? authorInitial,
        int? paragraphIndex,
        int? startRunIndex,
        int? endRunIndex,
        int? commentIndex,
        string? replyText)
    {
        return operation.ToLower() switch
        {
            "add" => BuildAddParameters(text, author, authorInitial, paragraphIndex, startRunIndex, endRunIndex),
            "delete" => BuildDeleteParameters(commentIndex),
            "reply" => BuildReplyParameters(commentIndex, replyText, text, author, authorInitial),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the add comment operation.
    /// </summary>
    /// <param name="text">The comment text content.</param>
    /// <param name="author">The comment author name.</param>
    /// <param name="authorInitial">The author initials.</param>
    /// <param name="paragraphIndex">The paragraph index (0-based).</param>
    /// <param name="startRunIndex">The start run index.</param>
    /// <param name="endRunIndex">The end run index.</param>
    /// <returns>OperationParameters configured for adding a comment.</returns>
    private static OperationParameters BuildAddParameters(string? text, string? author, string? authorInitial,
        int? paragraphIndex, int? startRunIndex, int? endRunIndex)
    {
        var parameters = new OperationParameters();
        if (text != null) parameters.Set("text", text);
        parameters.Set("author", author ?? "Comment Author");
        if (authorInitial != null) parameters.Set("authorInitial", authorInitial);
        if (paragraphIndex.HasValue) parameters.Set("paragraphIndex", paragraphIndex.Value);
        if (startRunIndex.HasValue) parameters.Set("startRunIndex", startRunIndex.Value);
        if (endRunIndex.HasValue) parameters.Set("endRunIndex", endRunIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the delete comment operation.
    /// </summary>
    /// <param name="commentIndex">The comment index (0-based).</param>
    /// <returns>OperationParameters configured for deleting a comment.</returns>
    private static OperationParameters BuildDeleteParameters(int? commentIndex)
    {
        var parameters = new OperationParameters();
        if (commentIndex.HasValue) parameters.Set("commentIndex", commentIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the reply to comment operation.
    /// </summary>
    /// <param name="commentIndex">The comment index (0-based).</param>
    /// <param name="replyText">The reply text content.</param>
    /// <param name="text">Fallback text content if replyText is null.</param>
    /// <param name="author">The reply author name.</param>
    /// <param name="authorInitial">The author initials.</param>
    /// <returns>OperationParameters configured for replying to a comment.</returns>
    private static OperationParameters BuildReplyParameters(int? commentIndex, string? replyText, string? text,
        string? author, string? authorInitial)
    {
        var parameters = new OperationParameters();
        if (commentIndex.HasValue) parameters.Set("commentIndex", commentIndex.Value);
        parameters.Set("replyText", replyText ?? text);
        parameters.Set("author", author ?? "Reply Author");
        if (authorInitial != null) parameters.Set("authorInitial", authorInitial);
        return parameters;
    }
}
