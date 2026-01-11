using System.ComponentModel;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Handlers.Word.Hyperlink;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for managing Word hyperlinks (add, edit, delete, get)
///     Merges: WordAddHyperlinkTool, WordEditHyperlinkTool, WordDeleteHyperlinkTool, WordGetHyperlinksTool
/// </summary>
[McpServerToolType]
public class WordHyperlinkTool
{
    /// <summary>
    ///     Handler registry for hyperlink operations
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
    ///     Initializes a new instance of the WordHyperlinkTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordHyperlinkTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = WordHyperlinkHandlerRegistry.Create();
    }

    /// <summary>
    ///     Executes a Word hyperlink operation (add, edit, delete, get).
    /// </summary>
    /// <param name="operation">The operation to perform: add, edit, delete, get.</param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="text">Display text for the hyperlink (for add).</param>
    /// <param name="url">URL or target address (for add/edit).</param>
    /// <param name="subAddress">Internal bookmark name for document navigation (for add/edit).</param>
    /// <param name="paragraphIndex">Paragraph index to insert hyperlink after (0-based, for add).</param>
    /// <param name="tooltip">Tooltip text (for add/edit).</param>
    /// <param name="hyperlinkIndex">Hyperlink index (0-based, for edit/delete).</param>
    /// <param name="displayText">New display text (for edit).</param>
    /// <param name="keepText">Keep display text when deleting hyperlink (default: false).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "word_hyperlink")]
    [Description(@"Manage Word hyperlinks. Supports 4 operations: add, edit, delete, get.

Usage examples:
- Add hyperlink: word_hyperlink(operation='add', path='doc.docx', text='Link', url='https://example.com', paragraphIndex=0)
- Edit hyperlink: word_hyperlink(operation='edit', path='doc.docx', hyperlinkIndex=0, url='https://newurl.com')
- Delete hyperlink: word_hyperlink(operation='delete', path='doc.docx', hyperlinkIndex=0)
- Get hyperlinks: word_hyperlink(operation='get', path='doc.docx')")]
    public string Execute(
        [Description(@"Operation to perform.
- 'add': Add a hyperlink (required params: path, text, url)
- 'edit': Edit a hyperlink (required params: path, hyperlinkIndex, url)
- 'delete': Delete a hyperlink (required params: path, hyperlinkIndex)
- 'get': Get all hyperlinks (required params: path)")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Display text for the hyperlink (required for add operation)")]
        string? text = null,
        [Description(
            "URL or target address (required for add operation unless subAddress is provided, optional for edit operation)")]
        string? url = null,
        [Description(
            "Internal bookmark name for document navigation (e.g., '_Toc123456'). Use with empty url for internal links. (optional, for add/edit operations)")]
        string? subAddress = null,
        [Description(
            "Paragraph index to insert hyperlink after (0-based, optional, for add operation). When specified, creates a NEW paragraph after the specified paragraph (does not insert into existing paragraph). Valid range: 0 to (total paragraphs - 1), or -1 for document start.")]
        int? paragraphIndex = null,
        [Description("Tooltip text (optional, for add/edit operations)")]
        string? tooltip = null,
        [Description("Hyperlink index (0-based, required for edit/delete operations)")]
        int? hyperlinkIndex = null,
        [Description("New display text (optional, for edit operation)")]
        string? displayText = null,
        [Description(
            "Keep display text when deleting hyperlink (unlink instead of remove, optional, default: false, for delete operation)")]
        bool keepText = false)
    {
        var parameters = BuildParameters(operation, text, url, subAddress, paragraphIndex, tooltip, hyperlinkIndex,
            displayText, keepText);

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
        string? text,
        string? url,
        string? subAddress,
        int? paragraphIndex,
        string? tooltip,
        int? hyperlinkIndex,
        string? displayText,
        bool keepText)
    {
        var parameters = new OperationParameters();

        switch (operation.ToLower())
        {
            case "add":
                if (text != null) parameters.Set("text", text);
                if (url != null) parameters.Set("url", url);
                if (subAddress != null) parameters.Set("subAddress", subAddress);
                if (paragraphIndex.HasValue) parameters.Set("paragraphIndex", paragraphIndex.Value);
                if (tooltip != null) parameters.Set("tooltip", tooltip);
                break;

            case "edit":
                parameters.Set("hyperlinkIndex", hyperlinkIndex ?? 0);
                if (url != null) parameters.Set("url", url);
                if (subAddress != null) parameters.Set("subAddress", subAddress);
                if (displayText != null) parameters.Set("displayText", displayText);
                if (tooltip != null) parameters.Set("tooltip", tooltip);
                break;

            case "delete":
                parameters.Set("hyperlinkIndex", hyperlinkIndex ?? 0);
                parameters.Set("keepText", keepText);
                break;
        }

        return parameters;
    }
}
