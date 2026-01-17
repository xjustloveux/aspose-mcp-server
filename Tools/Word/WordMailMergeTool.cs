using System.ComponentModel;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Tool for performing mail merge operations on Word document templates.
///     Dispatches operations to individual handlers via HandlerRegistry.
/// </summary>
[McpServerToolType]
public class WordMailMergeTool
{
    /// <summary>
    ///     Registry of mail merge operation handlers.
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
    ///     Initializes a new instance of the <see cref="WordMailMergeTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public WordMailMergeTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Word.MailMerge");
    }

    /// <summary>
    ///     Performs mail merge on a Word document template.
    /// </summary>
    /// <param name="operation">The operation to perform (currently only 'execute').</param>
    /// <param name="templatePath">Template file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID to use template from session.</param>
    /// <param name="outputPath">Output file path (required).</param>
    /// <param name="data">Key-value pairs for mail merge fields (for single record), as JSON object.</param>
    /// <param name="dataArray">Array of objects for multiple records, as JSON array.</param>
    /// <param name="cleanupOptions">Cleanup options to apply after mail merge, as comma-separated string.</param>
    /// <returns>A message indicating the mail merge result with field and file information.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when outputPath is not provided, neither templatePath nor sessionId is provided,
    ///     neither data nor dataArray is provided, or both data and dataArray are provided.
    /// </exception>
    /// <exception cref="InvalidOperationException">Thrown when session management is not enabled or document cloning fails.</exception>
    [McpServerTool(Name = "word_mail_merge")]
    [Description(@"Perform mail merge on a Word document template.

Usage examples:
- Single record: word_mail_merge(templatePath='template.docx', outputPath='output.docx', data={'name':'John','address':'123 Main St'})
- Multiple records: word_mail_merge(templatePath='template.docx', outputPath='output.docx', dataArray=[{'name':'John'},{'name':'Jane'}])
- From session: word_mail_merge(sessionId='sess_xxx', outputPath='output.docx', data={'name':'John'})")]
    public string Execute(
        [Description(@"Operation to perform.
- 'execute': Execute mail merge (required params: outputPath, and either data or dataArray)")]
        string operation = "execute",
        [Description("Template file path (required if no sessionId)")]
        string? templatePath = null,
        [Description("Session ID to use template from session")]
        string? sessionId = null,
        [Description(
            "Output file path (required). For multiple records, files will be named output_1.docx, output_2.docx, etc.")]
        string? outputPath = null,
        [Description("Key-value pairs for mail merge fields (for single record), as JSON object")]
        string? data = null,
        [Description(
            "Array of objects for multiple records, as JSON array. Each object contains key-value pairs for mail merge fields. Example: [{'name':'John','city':'NYC'},{'name':'Jane','city':'LA'}]")]
        string? dataArray = null,
        [Description(@"Cleanup options to apply after mail merge, as comma-separated string. Available options:
- 'removeUnusedFields': Remove merge fields that were not populated
- 'removeUnusedRegions': Remove mail merge regions that were not populated
- 'removeEmptyParagraphs': Remove paragraphs that become empty after merge
- 'removeContainingFields': Remove paragraphs containing empty merge fields
- 'removeStaticFields': Remove static fields (like PAGE, DATE)
Default: 'removeUnusedFields,removeEmptyParagraphs'")]
        string? cleanupOptions = null)
    {
        if (string.IsNullOrEmpty(templatePath) && string.IsNullOrEmpty(sessionId))
            throw new ArgumentException("Either templatePath or sessionId must be provided");

        if (!string.IsNullOrEmpty(templatePath))
            SecurityHelper.ValidateFilePath(templatePath, nameof(templatePath), true);

        var parameters = BuildParameters(outputPath, data, dataArray, cleanupOptions);

        var handler = _handlerRegistry.GetHandler(operation);

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, templatePath, _identityAccessor);

        var operationContext = new OperationContext<Document>
        {
            Document = ctx.Document,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = templatePath,
            OutputPath = outputPath
        };

        var result = handler.Execute(operationContext, parameters);

        return result;
    }

    /// <summary>
    ///     Builds the operation parameters from input values.
    /// </summary>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="data">The JSON object for single record mail merge.</param>
    /// <param name="dataArray">The JSON array for multiple records mail merge.</param>
    /// <param name="cleanupOptions">The cleanup options string.</param>
    /// <returns>The operation parameters.</returns>
    private static OperationParameters BuildParameters(string? outputPath, string? data, string? dataArray,
        string? cleanupOptions)
    {
        var parameters = new OperationParameters();
        parameters.Set("outputPath", outputPath);
        parameters.Set("data", data);
        parameters.Set("dataArray", dataArray);
        parameters.Set("cleanupOptions", cleanupOptions);
        return parameters;
    }
}
