using System.ComponentModel;
using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing attachments in PDF documents (add, delete, get)
/// </summary>
[McpServerToolType]
public class PdfAttachmentTool
{
    /// <summary>
    ///     Handler registry for attachment operations.
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
    ///     Initializes a new instance of the <see cref="PdfAttachmentTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public PdfAttachmentTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Pdf.Attachment");
    }

    /// <summary>
    ///     Executes a PDF attachment operation (add, delete, get).
    /// </summary>
    /// <param name="operation">The operation to perform: add, delete, get.</param>
    /// <param name="path">PDF file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="attachmentPath">Attachment file path (required for add).</param>
    /// <param name="attachmentName">Attachment name in PDF (required for add, delete).</param>
    /// <param name="description">Attachment description (optional for add).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "pdf_attachment")]
    [Description(@"Manage attachments in PDF documents. Supports 3 operations: add, delete, get.

Usage examples:
- Add attachment: pdf_attachment(operation='add', path='doc.pdf', attachmentPath='file.pdf', attachmentName='attachment.pdf')
- Delete attachment: pdf_attachment(operation='delete', path='doc.pdf', attachmentName='attachment.pdf')
- Get attachments: pdf_attachment(operation='get', path='doc.pdf')")]
    public string Execute( // NOSONAR S107 - MCP protocol requires multiple parameters
        [Description(@"Operation to perform.
- 'add': Add an attachment (required params: path, attachmentPath, attachmentName)
- 'delete': Delete an attachment (required params: path, attachmentName)
- 'get': Get all attachments (required params: path)")]
        string operation,
        [Description("PDF file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Attachment file path (required for add)")]
        string? attachmentPath = null,
        [Description("Attachment name in PDF (required for add, delete)")]
        string? attachmentName = null,
        [Description("Attachment description (optional for add)")]
        string? description = null)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, attachmentPath, attachmentName, description);

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
            return result;

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return $"{result}\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters( // NOSONAR S107 - MCP protocol parameter building
        string operation,
        string? attachmentPath,
        string? attachmentName,
        string? description)
    {
        return operation.ToLowerInvariant() switch
        {
            "add" => BuildAddParameters(attachmentPath, attachmentName, description),
            "delete" => BuildDeleteParameters(attachmentName),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the add attachment operation.
    /// </summary>
    /// <param name="attachmentPath">Path to the file to attach.</param>
    /// <param name="attachmentName">Name of the attachment in PDF.</param>
    /// <param name="description">Optional description for the attachment.</param>
    /// <returns>OperationParameters configured for adding an attachment.</returns>
    private static OperationParameters BuildAddParameters(string? attachmentPath, string? attachmentName,
        string? description)
    {
        var parameters = new OperationParameters();
        if (attachmentPath != null) parameters.Set("attachmentPath", attachmentPath);
        if (attachmentName != null) parameters.Set("attachmentName", attachmentName);
        if (description != null) parameters.Set("description", description);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the delete attachment operation.
    /// </summary>
    /// <param name="attachmentName">Name of the attachment to delete.</param>
    /// <returns>OperationParameters configured for deleting an attachment.</returns>
    private static OperationParameters BuildDeleteParameters(string? attachmentName)
    {
        var parameters = new OperationParameters();
        if (attachmentName != null) parameters.Set("attachmentName", attachmentName);
        return parameters;
    }
}
