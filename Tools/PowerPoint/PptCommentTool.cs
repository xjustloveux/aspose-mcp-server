using System.ComponentModel;
using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Tool for managing PowerPoint comments (add, get, delete, reply).
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.PowerPoint.Comment")]
[McpServerToolType]
public class PptCommentTool
{
    /// <summary>
    ///     Handler registry for comment operations.
    /// </summary>
    private readonly HandlerRegistry<Presentation> _handlerRegistry;

    /// <summary>
    ///     Identity accessor for session isolation.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document session handling.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PptCommentTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory editing.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public PptCommentTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry =
            HandlerRegistry<Presentation>.CreateFromNamespace("AsposeMcpServer.Handlers.PowerPoint.Comment");
    }

    /// <summary>
    ///     Executes a PowerPoint comment operation (add, get, delete, reply).
    /// </summary>
    /// <param name="operation">The operation to perform.</param>
    /// <param name="path">Presentation file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (optional).</param>
    /// <param name="slideIndex">Slide index (0-based).</param>
    /// <param name="commentIndex">Comment index (0-based, for delete/reply).</param>
    /// <param name="text">Comment text (required for add/reply).</param>
    /// <param name="author">Author name (required for add/reply).</param>
    /// <param name="x">X position in points (for add).</param>
    /// <param name="y">Y position in points (for add).</param>
    /// <returns>Operation result.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "ppt_comment",
        Title = "PowerPoint Comment Operations",
        Destructive = false,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(@"Manage PowerPoint comments. Supports 4 operations: add, get, delete, reply.

Usage examples:
- Add comment: ppt_comment(operation='add', path='file.pptx', text='Review needed', author='John')
- Get comments: ppt_comment(operation='get', path='file.pptx', slideIndex=0)
- Delete comment: ppt_comment(operation='delete', path='file.pptx', commentIndex=0)
- Reply to comment: ppt_comment(operation='reply', path='file.pptx', commentIndex=0, text='Done', author='Jane')")]
    public object Execute(
        [Description(@"Operation to perform.
- 'add': Add comment to slide (required: text, author)
- 'get': Get comments from slide
- 'delete': Delete comment (required: commentIndex)
- 'reply': Reply to comment (required: commentIndex, text, author)")]
        string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (optional)")]
        string? outputPath = null,
        [Description("Slide index (0-based, default: 0)")]
        int slideIndex = 0,
        [Description("Comment index (0-based, for delete/reply)")]
        int? commentIndex = null,
        [Description("Comment text (required for add/reply)")]
        string? text = null,
        [Description("Author name (required for add/reply)")]
        string? author = null,
        [Description("X position in points (for add, default: 0)")]
        float x = 0,
        [Description("Y position in points (for add, default: 0)")]
        float y = 0)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, slideIndex, commentIndex, text, author, x, y);
        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<Presentation>
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
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters(
        string operation, int slideIndex, int? commentIndex, string? text, string? author, float x, float y)
    {
        var parameters = new OperationParameters();
        parameters.Set("slideIndex", slideIndex);

        return operation.ToLowerInvariant() switch
        {
            "add" => BuildAddParameters(parameters, text, author, x, y),
            "get" => parameters,
            "delete" => BuildDeleteParameters(parameters, commentIndex),
            "reply" => BuildReplyParameters(parameters, commentIndex, text, author),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for the add comment operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters with slideIndex.</param>
    /// <param name="text">The comment text.</param>
    /// <param name="author">The author name.</param>
    /// <param name="x">The X position.</param>
    /// <param name="y">The Y position.</param>
    /// <returns>OperationParameters configured for adding a comment.</returns>
    private static OperationParameters BuildAddParameters(OperationParameters parameters, string? text, string? author,
        float x, float y)
    {
        if (text != null) parameters.Set("text", text);
        if (author != null) parameters.Set("author", author);
        parameters.Set("x", x);
        parameters.Set("y", y);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the delete comment operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters with slideIndex.</param>
    /// <param name="commentIndex">The comment index to delete.</param>
    /// <returns>OperationParameters configured for deleting a comment.</returns>
    private static OperationParameters BuildDeleteParameters(OperationParameters parameters, int? commentIndex)
    {
        if (commentIndex.HasValue) parameters.Set("commentIndex", commentIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the reply comment operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters with slideIndex.</param>
    /// <param name="commentIndex">The comment index to reply to.</param>
    /// <param name="text">The reply text.</param>
    /// <param name="author">The reply author name.</param>
    /// <returns>OperationParameters configured for replying to a comment.</returns>
    private static OperationParameters BuildReplyParameters(OperationParameters parameters, int? commentIndex,
        string? text, string? author)
    {
        if (commentIndex.HasValue) parameters.Set("commentIndex", commentIndex.Value);
        if (text != null) parameters.Set("text", text);
        if (author != null) parameters.Set("author", author);
        return parameters;
    }
}
