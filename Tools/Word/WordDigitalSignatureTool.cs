using System.ComponentModel;
using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Tool for managing digital signatures in Word documents (sign, verify, remove, list).
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Word.DigitalSignature")]
[McpServerToolType]
public class WordDigitalSignatureTool
{
    /// <summary>
    ///     Handler registry for digital signature operations.
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
    ///     Initializes a new instance of the <see cref="WordDigitalSignatureTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public WordDigitalSignatureTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry =
            HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Word.DigitalSignature");
    }

    /// <summary>
    ///     Executes a digital signature operation on a Word document (sign, verify, remove, list).
    /// </summary>
    /// <param name="operation">The operation to perform: sign, verify, remove, list.</param>
    /// <param name="path">Word document file path (required for all operations).</param>
    /// <param name="outputPath">Output file path (required for sign, remove).</param>
    /// <param name="certificatePath">PFX certificate file path (required for sign).</param>
    /// <param name="certificatePassword">Certificate password (required for sign).</param>
    /// <param name="comments">Comments for the signature (optional, for sign).</param>
    /// <returns>A message or data indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the certificate file is not found.</exception>
    [McpServerTool(
        Name = "word_digital_signature",
        Title = "Word Digital Signature Operations",
        Destructive = true,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(@"Manage digital signatures in Word documents. Supports 4 operations: sign, verify, remove, list.

Usage examples:
- Sign document: word_digital_signature(operation='sign', path='doc.docx', outputPath='signed.docx', certificatePath='cert.pfx', certificatePassword='pass')
- Verify signatures: word_digital_signature(operation='verify', path='signed.docx')
- Remove signatures: word_digital_signature(operation='remove', path='signed.docx', outputPath='unsigned.docx')
- List signatures: word_digital_signature(operation='list', path='signed.docx')")]
    public object Execute(
        [Description(@"Operation to perform.
- 'sign': Sign document with digital certificate (required params: path, outputPath, certificatePath, certificatePassword)
- 'verify': Verify all digital signatures (required params: path)
- 'remove': Remove all digital signatures (required params: path, outputPath)
- 'list': List all digital signatures (required params: path)")]
        string operation,
        [Description("Word document file path (required for all operations)")]
        string? path = null,
        [Description("Output file path (required for sign, remove)")]
        string? outputPath = null,
        [Description("PFX certificate file path (required for sign)")]
        string? certificatePath = null,
        [Description("Certificate password (required for sign)")]
        string? certificatePassword = null,
        [Description("Comments for the signature (optional, for sign)")]
        string? comments = null)
    {
        var parameters = BuildParameters(operation, path, outputPath, certificatePath, certificatePassword, comments);

        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<Document>
        {
            Document = null!,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SourcePath = path,
            OutputPath = outputPath
        };

        var result = handler.Execute(operationContext, parameters);
        return ResultHelper.FinalizeResult((dynamic)result, outputPath, (string?)null);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters(
        string operation,
        string? path,
        string? outputPath,
        string? certificatePath,
        string? certificatePassword,
        string? comments)
    {
        return operation.ToLowerInvariant() switch
        {
            "sign" => BuildSignParameters(path, outputPath, certificatePath, certificatePassword, comments),
            "verify" => BuildPathOnlyParameters(path),
            "remove" => BuildRemoveParameters(path, outputPath),
            "list" => BuildPathOnlyParameters(path),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the sign operation.
    /// </summary>
    /// <param name="path">The source document file path.</param>
    /// <param name="outputPath">The destination file path.</param>
    /// <param name="certificatePath">The PFX certificate file path.</param>
    /// <param name="certificatePassword">The certificate password.</param>
    /// <param name="comments">Optional signature comments.</param>
    /// <returns>OperationParameters configured for signing.</returns>
    private static OperationParameters BuildSignParameters(string? path, string? outputPath,
        string? certificatePath, string? certificatePassword, string? comments)
    {
        var parameters = new OperationParameters();
        if (path != null) parameters.Set("path", path);
        if (outputPath != null) parameters.Set("outputPath", outputPath);
        if (certificatePath != null) parameters.Set("certificatePath", certificatePath);
        if (certificatePassword != null) parameters.Set("certificatePassword", certificatePassword);
        if (comments != null) parameters.Set("comments", comments);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for operations that only need a path.
    /// </summary>
    /// <param name="path">The document file path.</param>
    /// <returns>OperationParameters with path only.</returns>
    private static OperationParameters BuildPathOnlyParameters(string? path)
    {
        var parameters = new OperationParameters();
        if (path != null) parameters.Set("path", path);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the remove operation.
    /// </summary>
    /// <param name="path">The source document file path.</param>
    /// <param name="outputPath">The destination file path.</param>
    /// <returns>OperationParameters configured for removing signatures.</returns>
    private static OperationParameters BuildRemoveParameters(string? path, string? outputPath)
    {
        var parameters = new OperationParameters();
        if (path != null) parameters.Set("path", path);
        if (outputPath != null) parameters.Set("outputPath", outputPath);
        return parameters;
    }
}
