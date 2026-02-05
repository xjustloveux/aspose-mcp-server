using System.ComponentModel;
using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for validating and converting PDF documents for compliance standards (PDF/A, PDF/UA)
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Pdf.Compliance")]
[McpServerToolType]
public class PdfComplianceTool
{
    /// <summary>
    ///     Handler registry for compliance operations.
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
    ///     Initializes a new instance of the <see cref="PdfComplianceTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public PdfComplianceTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Pdf.Compliance");
    }

    /// <summary>
    ///     Executes a PDF compliance operation (validate or convert).
    /// </summary>
    /// <param name="operation">The operation to perform: validate, convert.</param>
    /// <param name="path">PDF file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="format">Compliance format (e.g., pdf/a-1b, pdf/a-2a, pdf/ua-1).</param>
    /// <param name="logPath">Path to write validation or conversion log file.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "pdf_compliance",
        Title = "PDF Compliance Operations",
        Destructive = false,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(
        @"Validate and convert PDF documents for compliance standards (PDF/A, PDF/UA). Supports 2 operations: validate, convert.

Usage examples:
- Validate PDF/A-1b: pdf_compliance(operation='validate', path='doc.pdf', format='pdf/a-1b')
- Validate with log: pdf_compliance(operation='validate', path='doc.pdf', format='pdf/a-2b', logPath='validation.log')
- Convert to PDF/A: pdf_compliance(operation='convert', path='doc.pdf', outputPath='compliant.pdf', format='pdf/a-1b')
- Convert to PDF/UA: pdf_compliance(operation='convert', path='doc.pdf', outputPath='accessible.pdf', format='pdf/ua-1')

Supported formats: pdf/a-1a, pdf/a-1b, pdf/a-2a, pdf/a-2b, pdf/a-3a, pdf/a-3b, pdf/ua-1")]
    public object Execute(
        [Description("Operation: validate, convert")]
        string operation,
        [Description("PDF file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Compliance format (e.g., pdf/a-1b, pdf/a-2a, pdf/ua-1)")]
        string format = "pdf/a-1b",
        [Description("Path to write validation or conversion log file")]
        string? logPath = null)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, format, logPath);

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

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return ResultHelper.FinalizeResult((dynamic)result, ctx, outputPath);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <param name="operation">The operation to perform.</param>
    /// <param name="format">The compliance format.</param>
    /// <param name="logPath">The optional log file path.</param>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters(
        string operation,
        string format,
        string? logPath)
    {
        return operation.ToLowerInvariant() switch
        {
            "validate" => BuildComplianceParameters(format, logPath),
            "convert" => BuildComplianceParameters(format, logPath),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for compliance operations (validate and convert).
    /// </summary>
    /// <param name="format">The compliance format.</param>
    /// <param name="logPath">The optional log file path.</param>
    /// <returns>OperationParameters configured for compliance operations.</returns>
    private static OperationParameters BuildComplianceParameters(string format, string? logPath)
    {
        var parameters = new OperationParameters();
        parameters.Set("format", format);
        if (logPath != null) parameters.Set("logPath", logPath);
        return parameters;
    }
}
