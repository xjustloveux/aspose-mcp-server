using System.ComponentModel;
using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Tool for managing revision tracking in Word documents
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Word.Revision")]
[McpServerToolType]
public class WordRevisionTool
{
    /// <summary>
    ///     Handler registry for revision operations
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
    ///     Initializes a new instance of the WordRevisionTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordRevisionTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Word.Revision");
    }

    /// <summary>
    ///     Executes a Word revision operation (get_revisions, accept_all, reject_all, manage, compare).
    /// </summary>
    /// <param name="operation">The operation to perform: get_revisions, accept_all, reject_all, manage, compare.</param>
    /// <param name="path">Word document file path (required if no sessionId for most operations).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (optional for most operations, required for compare).</param>
    /// <param name="revisionIndex">Revision index (0-based, required for manage operation).</param>
    /// <param name="action">Action for manage operation: accept, reject (default: accept).</param>
    /// <param name="originalPath">Original document file path (required for compare).</param>
    /// <param name="revisedPath">Revised document file path (required for compare).</param>
    /// <param name="authorName">Author name for revisions (for compare, default: 'Comparison').</param>
    /// <param name="ignoreFormatting">Ignore formatting changes in comparison (for compare, default: false).</param>
    /// <param name="ignoreComments">Ignore comments in comparison (for compare, default: false).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get_revisions.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "word_revision",
        Title = "Word Revision Operations",
        Destructive = true,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(
        @"Manage revisions in Word documents. Supports 5 operations: get_revisions, accept_all, reject_all, manage, compare.

Usage examples:
- Get revisions: word_revision(operation='get_revisions', path='doc.docx')
- Accept all: word_revision(operation='accept_all', path='doc.docx')
- Reject all: word_revision(operation='reject_all', path='doc.docx')
- Manage specific revision: word_revision(operation='manage', path='doc.docx', revisionIndex=0, action='accept')
- Compare documents: word_revision(operation='compare', path='output.docx', originalPath='original.docx', revisedPath='revised.docx', ignoreFormatting=true)

Notes:
- The 'manage' operation accepts or rejects a specific revision by index (0-based)
- Use 'get_revisions' first to see all revisions and their indices
- Compare operation can optionally ignore formatting and comments changes")]
    public object Execute(
        [Description("Operation: get_revisions, accept_all, reject_all, manage, compare")]
        string operation,
        [Description("Document file path (required if no sessionId for most operations)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (optional for most operations, required for compare)")]
        string? outputPath = null,
        [Description("Revision index (0-based, required for manage operation)")]
        int? revisionIndex = null,
        [Description("Action for manage operation: accept, reject (default: accept)")]
        string action = "accept",
        [Description("Original document file path (required for compare)")]
        string? originalPath = null,
        [Description("Revised document file path (required for compare)")]
        string? revisedPath = null,
        [Description("Author name for revisions (for compare, default: 'Comparison')")]
        string authorName = "Comparison",
        [Description("Ignore formatting changes in comparison (for compare, default: false)")]
        bool ignoreFormatting = false,
        [Description("Ignore comments in comparison (for compare, default: false)")]
        bool ignoreComments = false)
    {
        var normalizedOperation = operation.ToLower();

        // Compare operation doesn't use session/document context - it loads its own files
        if (normalizedOperation == "compare")
        {
            var compareParameters = new OperationParameters();
            if (outputPath != null) compareParameters.Set("outputPath", outputPath);
            if (originalPath != null) compareParameters.Set("originalPath", originalPath);
            if (revisedPath != null) compareParameters.Set("revisedPath", revisedPath);
            compareParameters.Set("authorName", authorName);
            compareParameters.Set("ignoreFormatting", ignoreFormatting);
            compareParameters.Set("ignoreComments", ignoreComments);

            var compareHandler = _handlerRegistry.GetHandler(normalizedOperation);

            // CompareDocumentsHandler doesn't need a document context
            var dummyContext = new OperationContext<Document>
            {
                Document = null!,
                SessionManager = _sessionManager,
                IdentityAccessor = _identityAccessor,
                SessionId = null,
                SourcePath = null,
                OutputPath = outputPath
            };

            var compareResult = compareHandler.Execute(dummyContext, compareParameters);
            return ResultHelper.FinalizeResult((dynamic)compareResult, outputPath, sessionId);
        }

        var parameters = BuildParameters(normalizedOperation, revisionIndex, action);

        var handler = _handlerRegistry.GetHandler(normalizedOperation);

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

        return ResultHelper.FinalizeResult((dynamic)result, ctx, effectiveOutputPath);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters(
        string operation,
        int? revisionIndex,
        string action)
    {
        return operation switch
        {
            "get_revisions" or "accept_all" or "reject_all" => new OperationParameters(),
            "manage" => BuildManageParameters(revisionIndex, action),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the manage operation.
    /// </summary>
    /// <param name="revisionIndex">The revision index (0-based).</param>
    /// <param name="action">The action to perform: accept, reject.</param>
    /// <returns>OperationParameters configured for managing a specific revision.</returns>
    private static OperationParameters BuildManageParameters(int? revisionIndex, string action)
    {
        var parameters = new OperationParameters();
        if (revisionIndex.HasValue) parameters.Set("revisionIndex", revisionIndex.Value);
        parameters.Set("action", action);
        return parameters;
    }
}
