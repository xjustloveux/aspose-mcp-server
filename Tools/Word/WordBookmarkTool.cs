using System.ComponentModel;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for managing Word bookmarks (add, edit, delete, get, goto)
/// </summary>
[McpServerToolType]
public class WordBookmarkTool
{
    /// <summary>
    ///     Handler registry for bookmark operations
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
    ///     Initializes a new instance of the WordBookmarkTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordBookmarkTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Word.Bookmark");
    }

    /// <summary>
    ///     Executes a Word bookmark operation (add, edit, delete, get, goto).
    /// </summary>
    /// <param name="operation">The operation to perform: add, edit, delete, get, goto.</param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="name">Bookmark name.</param>
    /// <param name="text">Text content for bookmark.</param>
    /// <param name="paragraphIndex">Paragraph index (0-based, -1 for beginning).</param>
    /// <param name="newName">New bookmark name (for edit).</param>
    /// <param name="newText">New text content (for edit).</param>
    /// <param name="keepText">Keep text when deleting (default: true).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "word_bookmark")]
    [Description(@"Manage Word bookmarks. Supports 5 operations: add, edit, delete, get, goto.

Usage examples:
- Add bookmark: word_bookmark(operation='add', path='doc.docx', name='bookmark1', text='Bookmarked text')
- Edit bookmark: word_bookmark(operation='edit', path='doc.docx', name='bookmark1', text='Updated text')
- Delete bookmark: word_bookmark(operation='delete', path='doc.docx', name='bookmark1')
- Get bookmarks: word_bookmark(operation='get', path='doc.docx')
- Goto bookmark: word_bookmark(operation='goto', path='doc.docx', name='bookmark1')")]
    public string Execute( // NOSONAR S107 - MCP protocol requires multiple parameters
        [Description("Operation: add, edit, delete, get, goto")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Bookmark name")] string? name = null,
        [Description("Text content for bookmark")]
        string? text = null,
        [Description("Paragraph index (0-based, -1 for beginning)")]
        int? paragraphIndex = null,
        [Description("New bookmark name (for edit)")]
        string? newName = null,
        [Description("New text content (for edit)")]
        string? newText = null,
        [Description("Keep text when deleting (default: true)")]
        bool keepText = true)
    {
        var parameters = BuildParameters(operation, name, text, paragraphIndex, newName, newText, keepText);

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
        string? name,
        string? text,
        int? paragraphIndex,
        string? newName,
        string? newText,
        bool keepText)
    {
        return operation.ToLower() switch
        {
            "add" => BuildAddParameters(name, text, paragraphIndex),
            "edit" => BuildEditParameters(name, newName, newText, text),
            "delete" => BuildDeleteParameters(name, keepText),
            "goto" or "get" => BuildGotoParameters(name),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the add operation.
    /// </summary>
    /// <param name="name">The bookmark name.</param>
    /// <param name="text">The text content for the bookmark.</param>
    /// <param name="paragraphIndex">The paragraph index (0-based, -1 for beginning).</param>
    /// <returns>OperationParameters configured for adding a bookmark.</returns>
    private static OperationParameters BuildAddParameters(string? name, string? text, int? paragraphIndex)
    {
        var parameters = new OperationParameters();
        if (name != null) parameters.Set("name", name);
        if (text != null) parameters.Set("text", text);
        if (paragraphIndex.HasValue) parameters.Set("paragraphIndex", paragraphIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the edit operation.
    /// </summary>
    /// <param name="name">The bookmark name to edit.</param>
    /// <param name="newName">The new bookmark name.</param>
    /// <param name="newText">The new text content.</param>
    /// <param name="text">The text content (alternative to newText).</param>
    /// <returns>OperationParameters configured for editing a bookmark.</returns>
    private static OperationParameters BuildEditParameters(string? name, string? newName, string? newText, string? text)
    {
        var parameters = new OperationParameters();
        if (name != null) parameters.Set("name", name);
        if (newName != null) parameters.Set("newName", newName);
        if (newText != null) parameters.Set("newText", newText);
        else if (text != null) parameters.Set("newText", text);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the delete operation.
    /// </summary>
    /// <param name="name">The bookmark name to delete.</param>
    /// <param name="keepText">Whether to keep the text when deleting.</param>
    /// <returns>OperationParameters configured for deleting a bookmark.</returns>
    private static OperationParameters BuildDeleteParameters(string? name, bool keepText)
    {
        var parameters = new OperationParameters();
        if (name != null) parameters.Set("name", name);
        parameters.Set("keepText", keepText);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the goto or get operation.
    /// </summary>
    /// <param name="name">The bookmark name to navigate to or get.</param>
    /// <returns>OperationParameters configured for goto or get bookmark.</returns>
    private static OperationParameters BuildGotoParameters(string? name)
    {
        var parameters = new OperationParameters();
        if (name != null) parameters.Set("name", name);
        return parameters;
    }
}
