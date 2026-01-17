using System.ComponentModel;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for managing Word document protection (protect, unprotect)
///     Merges: WordProtectTool, WordUnprotectTool
/// </summary>
[McpServerToolType]
public class WordProtectionTool
{
    /// <summary>
    ///     Handler registry for protection operations
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
    ///     Initializes a new instance of the WordProtectionTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordProtectionTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Word.Protection");
    }

    /// <summary>
    ///     Executes a Word protection operation (protect, unprotect).
    /// </summary>
    /// <param name="operation">The operation to perform: protect, unprotect.</param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (if not provided, overwrites input).</param>
    /// <param name="password">Protection password (required for protect, optional for unprotect).</param>
    /// <param name="protectionType">Protection type: ReadOnly, AllowOnlyComments, AllowOnlyFormFields, AllowOnlyRevisions.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when password is missing for protect operation or the operation is unknown.</exception>
    [McpServerTool(Name = "word_protection")]
    [Description(@"Protect or unprotect a Word document. Supports 2 operations: protect, unprotect.

Usage examples:
- Protect document: word_protection(operation='protect', path='doc.docx', password='password', protectionType='ReadOnly')
- Unprotect document: word_protection(operation='unprotect', path='doc.docx', password='password')

Protection types:
- ReadOnly: Prevent all modifications (most restrictive)
- AllowOnlyComments: Allow only adding comments
- AllowOnlyFormFields: Allow only filling in form fields
- AllowOnlyRevisions: Allow only tracked changes

Notes:
- Password is required for 'protect' operation (cannot be empty)
- Password is optional for 'unprotect' (some documents may not require password)
- If unprotect fails, verify the password is correct
- For encrypted documents (with open password), the same password will be used to open the file")]
    public string Execute(
        [Description("Operation: protect, unprotect")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (if not provided, overwrites input)")]
        string? outputPath = null,
        [Description("Protection password (required for protect operation, optional for unprotect)")]
        string? password = null,
        [Description(
            "Protection type: 'ReadOnly', 'AllowOnlyComments', 'AllowOnlyFormFields', 'AllowOnlyRevisions' (required for protect operation)")]
        string protectionType = "ReadOnly")
    {
        var parameters = BuildParameters(operation, password, protectionType);

        var handler = _handlerRegistry.GetHandler(operation);

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor, password);

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

        var shouldSave = operationContext.IsModified || (outputPath != null && outputPath != path);
        if (shouldSave)
            ctx.Save(effectiveOutputPath);

        return ctx.IsSession ? result : $"{result}\n{ctx.GetOutputMessage(effectiveOutputPath)}";
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters(
        string operation,
        string? password,
        string protectionType)
    {
        return operation.ToLower() switch
        {
            "protect" => BuildProtectParameters(password, protectionType),
            "unprotect" => BuildUnprotectParameters(password),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the protect document operation.
    /// </summary>
    /// <param name="password">The protection password.</param>
    /// <param name="protectionType">The protection type (ReadOnly, AllowOnlyComments, etc.).</param>
    /// <returns>OperationParameters configured for protecting a document.</returns>
    private static OperationParameters BuildProtectParameters(string? password, string protectionType)
    {
        var parameters = new OperationParameters();
        if (password != null) parameters.Set("password", password);
        parameters.Set("protectionType", protectionType);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the unprotect document operation.
    /// </summary>
    /// <param name="password">The password to unprotect (optional).</param>
    /// <returns>OperationParameters configured for unprotecting a document.</returns>
    private static OperationParameters BuildUnprotectParameters(string? password)
    {
        var parameters = new OperationParameters();
        if (password != null) parameters.Set("password", password);
        return parameters;
    }
}
