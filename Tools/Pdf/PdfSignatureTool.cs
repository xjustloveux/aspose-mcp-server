using System.ComponentModel;
using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing digital signatures in PDF documents (sign, delete, get)
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Pdf.Signature")]
[McpServerToolType]
public class PdfSignatureTool
{
    /// <summary>
    ///     Handler registry for signature operations.
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
    ///     Initializes a new instance of the <see cref="PdfSignatureTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public PdfSignatureTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Pdf.Signature");
    }

    /// <summary>
    ///     Executes a PDF signature operation (sign, delete, get).
    /// </summary>
    /// <param name="operation">The operation to perform: sign, delete, get.</param>
    /// <param name="path">PDF file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="certificatePath">Path to certificate file (.pfx, required for sign).</param>
    /// <param name="certificatePassword">Certificate password (required for sign).</param>
    /// <param name="reason">Reason for signing (for sign, optional).</param>
    /// <param name="location">Location of signing (for sign, optional).</param>
    /// <param name="signatureIndex">Signature index (0-based, required for delete).</param>
    /// <param name="pageIndex">Page index to place signature (1-based, for sign, default: 1).</param>
    /// <param name="x">X position of signature in PDF coordinates (for sign).</param>
    /// <param name="y">Y position of signature in PDF coordinates (for sign).</param>
    /// <param name="width">Width of signature rectangle in PDF points (for sign).</param>
    /// <param name="height">Height of signature rectangle in PDF points (for sign).</param>
    /// <param name="imagePath">Path to signature appearance image (for sign, optional).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "pdf_signature",
        Title = "PDF Signature Operations",
        Destructive = true,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(@"Manage digital signatures in PDF documents. Supports 3 operations: sign, delete, get.

Usage examples:
- Sign PDF: pdf_signature(operation='sign', path='doc.pdf', certificatePath='cert.pfx', certificatePassword='password')
- Sign with position: pdf_signature(operation='sign', path='doc.pdf', certificatePath='cert.pfx', certificatePassword='password', pageIndex=1, x=100, y=100, width=200, height=100)
- Sign with image: pdf_signature(operation='sign', path='doc.pdf', certificatePath='cert.pfx', certificatePassword='password', imagePath='stamp.png')
- Delete signature: pdf_signature(operation='delete', path='doc.pdf', signatureIndex=0)
- Get signatures: pdf_signature(operation='get', path='doc.pdf')")]
    public object Execute(
        [Description("Operation: sign, delete, get")]
        string operation,
        [Description("PDF file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Path to certificate file (.pfx, required for sign)")]
        string? certificatePath = null,
        [Description("Certificate password (required for sign)")]
        string? certificatePassword = null,
        [Description("Reason for signing (for sign, optional)")]
        string reason = "Document approval",
        [Description("Location of signing (for sign, optional)")]
        string location = "",
        [Description("Signature index (0-based, required for delete)")]
        int signatureIndex = 0,
        [Description("Page index to place signature (1-based, for sign, default: 1)")]
        int pageIndex = 1,
        [Description("X position of signature in PDF coordinates (for sign, default: 100)")]
        int x = 100,
        [Description("Y position of signature in PDF coordinates (for sign, default: 100)")]
        int y = 100,
        [Description("Width of signature rectangle in PDF points (for sign, default: 200)")]
        int width = 200,
        [Description("Height of signature rectangle in PDF points (for sign, default: 100)")]
        int height = 100,
        [Description("Path to signature appearance image (for sign, optional)")]
        string? imagePath = null)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, certificatePath, certificatePassword, reason, location,
            signatureIndex, pageIndex, x, y, width, height, imagePath);

        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<Document>
        {
            Document = ctx.Document,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = path,
            OutputPath = outputPath
        };

        var result = handler.Execute(operationContext, parameters);

        if (string.Equals(operation, "get", StringComparison.OrdinalIgnoreCase))
            return ResultHelper.FinalizeResult((dynamic)result, ctx, outputPath);

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
        string operation,
        string? certificatePath,
        string? certificatePassword,
        string reason,
        string location,
        int signatureIndex,
        int pageIndex,
        int x,
        int y,
        int width,
        int height,
        string? imagePath)
    {
        return operation.ToLowerInvariant() switch
        {
            "sign" => BuildSignParameters(certificatePath, certificatePassword, reason, location, pageIndex, x, y,
                width, height, imagePath),
            "delete" => BuildDeleteParameters(signatureIndex),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the sign PDF operation.
    /// </summary>
    /// <param name="certificatePath">Path to the certificate file (.pfx).</param>
    /// <param name="certificatePassword">The certificate password.</param>
    /// <param name="reason">The reason for signing.</param>
    /// <param name="location">The location of signing.</param>
    /// <param name="pageIndex">The page index to place signature (1-based).</param>
    /// <param name="x">X position of signature.</param>
    /// <param name="y">Y position of signature.</param>
    /// <param name="width">Width of signature rectangle.</param>
    /// <param name="height">Height of signature rectangle.</param>
    /// <param name="imagePath">Path to signature appearance image.</param>
    /// <returns>OperationParameters configured for signing a PDF.</returns>
    private static OperationParameters BuildSignParameters(string? certificatePath, string? certificatePassword,
        string reason, string location, int pageIndex, int x, int y, int width, int height, string? imagePath)
    {
        var parameters = new OperationParameters();
        if (certificatePath != null) parameters.Set("certificatePath", certificatePath);
        if (certificatePassword != null) parameters.Set("password", certificatePassword);
        parameters.Set("reason", reason);
        parameters.Set("location", location);
        parameters.Set("pageIndex", pageIndex);
        parameters.Set("x", x);
        parameters.Set("y", y);
        parameters.Set("width", width);
        parameters.Set("height", height);
        if (imagePath != null) parameters.Set("imagePath", imagePath);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the delete signature operation.
    /// </summary>
    /// <param name="signatureIndex">The signature index to delete (0-based).</param>
    /// <returns>OperationParameters configured for deleting a signature.</returns>
    private static OperationParameters BuildDeleteParameters(int signatureIndex)
    {
        var parameters = new OperationParameters();
        parameters.Set("signatureIndex", signatureIndex);
        return parameters;
    }
}
