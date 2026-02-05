using System.ComponentModel;
using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint security operations
///     (encrypt, decrypt, set_write_protection, remove_write_protection, mark_final, get_status).
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.PowerPoint.Security")]
[McpServerToolType]
public class PptSecurityTool
{
    /// <summary>
    ///     Handler registry for security operations.
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
    ///     Initializes a new instance of the <see cref="PptSecurityTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory editing.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public PptSecurityTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry =
            HandlerRegistry<Presentation>.CreateFromNamespace("AsposeMcpServer.Handlers.PowerPoint.Security");
    }

    /// <summary>
    ///     Executes a PowerPoint security operation (encrypt, decrypt, set_write_protection,
    ///     remove_write_protection, mark_final, get_status).
    /// </summary>
    /// <param name="operation">The operation to perform.</param>
    /// <param name="path">Presentation file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (optional).</param>
    /// <param name="password">Password (required for encrypt, set_write_protection).</param>
    /// <param name="markAsFinal">Mark as final flag (for mark_final, default: true).</param>
    /// <returns>A message indicating the result of the operation, or security status data for get_status.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "ppt_security",
        Title = "PowerPoint Security Operations",
        Destructive = false,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(
        @"Manage PowerPoint security. Supports 6 operations: encrypt, decrypt, set_write_protection, remove_write_protection, mark_final, get_status.

Warning: If outputPath is not provided for write operations, the original file will be overwritten.

Usage examples:
- Encrypt: ppt_security(operation='encrypt', path='file.pptx', password='secret')
- Decrypt: ppt_security(operation='decrypt', path='file.pptx')
- Set write protection: ppt_security(operation='set_write_protection', path='file.pptx', password='edit_pass')
- Remove write protection: ppt_security(operation='remove_write_protection', path='file.pptx')
- Mark as final: ppt_security(operation='mark_final', path='file.pptx')
- Get status: ppt_security(operation='get_status', path='file.pptx')")]
    public object Execute(
        [Description(@"Operation to perform.
- 'encrypt': Encrypt presentation (required: password)
- 'decrypt': Remove encryption
- 'set_write_protection': Set write protection (required: password)
- 'remove_write_protection': Remove write protection
- 'mark_final': Mark presentation as final
- 'get_status': Get security status")]
        string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (optional)")]
        string? outputPath = null,
        [Description("Password (required for encrypt, set_write_protection)")]
        string? password = null,
        [Description("Mark as final flag (for mark_final, default: true)")]
        bool markAsFinal = true)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, password, markAsFinal);
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

        if (string.Equals(operation, "get_status", StringComparison.OrdinalIgnoreCase))
            return ResultHelper.FinalizeResult((dynamic)result, ctx, outputPath);

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return ResultHelper.FinalizeResult((dynamic)result, ctx, outputPath);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters based on the operation type.
    /// </summary>
    /// <param name="operation">The operation to perform.</param>
    /// <param name="password">The password for encrypt or set_write_protection operations.</param>
    /// <param name="markAsFinal">The mark-as-final flag for the mark_final operation.</param>
    /// <returns>OperationParameters configured for the specified operation.</returns>
    private static OperationParameters BuildParameters(string operation, string? password, bool markAsFinal)
    {
        var parameters = new OperationParameters();

        return operation.ToLowerInvariant() switch
        {
            "encrypt" or "set_write_protection" => BuildPasswordParameters(parameters, password),
            "mark_final" => BuildMarkFinalParameters(parameters, markAsFinal),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for operations that require a password.
    /// </summary>
    /// <param name="parameters">The base parameters to configure.</param>
    /// <param name="password">The password to set.</param>
    /// <returns>OperationParameters configured with the password.</returns>
    private static OperationParameters BuildPasswordParameters(OperationParameters parameters, string? password)
    {
        if (password != null) parameters.Set("password", password);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the mark_final operation.
    /// </summary>
    /// <param name="parameters">The base parameters to configure.</param>
    /// <param name="markAsFinal">Whether to mark or unmark the presentation as final.</param>
    /// <returns>OperationParameters configured with the markAsFinal flag.</returns>
    private static OperationParameters BuildMarkFinalParameters(OperationParameters parameters, bool markAsFinal)
    {
        parameters.Set("markAsFinal", markAsFinal);
        return parameters;
    }
}
