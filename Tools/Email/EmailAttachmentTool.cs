using System.ComponentModel;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Email;

/// <summary>
///     Tool for managing email attachments (list, add, remove, extract).
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Email.Attachment")]
[McpServerToolType]
public class EmailAttachmentTool
{
    /// <summary>
    ///     Handler registry for email attachment operations.
    /// </summary>
    private readonly HandlerRegistry<object> _handlerRegistry;

    /// <summary>
    ///     Initializes a new instance of the <see cref="EmailAttachmentTool" /> class.
    /// </summary>
    public EmailAttachmentTool()
    {
        _handlerRegistry =
            HandlerRegistry<object>.CreateFromNamespace("AsposeMcpServer.Handlers.Email.Attachment");
    }

    /// <summary>
    ///     Executes an email attachment operation (list, add, remove, extract, extract_all).
    /// </summary>
    /// <param name="operation">The operation to perform: list, add, remove, extract, extract_all.</param>
    /// <param name="path">Input email file path (.eml or .msg).</param>
    /// <param name="outputPath">Output email file path (required for add, remove).</param>
    /// <param name="attachmentPath">Path of the file to attach (required for add).</param>
    /// <param name="outputDir">Output directory for extracted attachments (required for extract, extract_all).</param>
    /// <param name="index">Zero-based attachment index (required for remove, extract).</param>
    /// <returns>Operation result depending on the operation type.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "email_attachment",
        Title = "Email Attachment Operations",
        Destructive = true,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(@"Manage email attachments. Supports 5 operations: list, add, remove, extract, extract_all.

Usage examples:
- List attachments: email_attachment(operation='list', path='email.eml')
- Add attachment: email_attachment(operation='add', path='email.eml', outputPath='out.eml', attachmentPath='file.pdf')
- Remove attachment: email_attachment(operation='remove', path='email.eml', outputPath='out.eml', index=0)
- Extract one: email_attachment(operation='extract', path='email.eml', outputDir='./attachments', index=0)
- Extract all: email_attachment(operation='extract_all', path='email.eml', outputDir='./attachments')

Supported email formats: EML, MSG
Supported output formats: EML, MSG, MHT/MHTML, HTML")]
    public object Execute(
        [Description(@"Operation to perform.
- 'list': List all attachments in an email (required params: path)
- 'add': Add a file as attachment (required params: path, outputPath, attachmentPath)
- 'remove': Remove an attachment by index (required params: path, outputPath, index)
- 'extract': Extract a specific attachment to a directory (required params: path, outputDir, index)
- 'extract_all': Extract all attachments to a directory (required params: path, outputDir)")]
        string operation,
        [Description("Input email file path (.eml or .msg)")]
        string path,
        [Description("Output email file path (required for add, remove)")]
        string? outputPath = null,
        [Description("Path of the file to attach (required for add)")]
        string? attachmentPath = null,
        [Description("Output directory for extracted attachments (required for extract, extract_all)")]
        string? outputDir = null,
        [Description("Zero-based attachment index (required for remove, extract)")]
        int? index = null)
    {
        var parameters = BuildParameters(path, outputPath, attachmentPath, outputDir, index);
        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<object>
        {
            Document = new object(),
            SourcePath = path,
            OutputPath = outputPath
        };

        var result = handler.Execute(operationContext, parameters);

        var effectiveOutputPath = operation switch
        {
            "extract" or "extract_all" => outputDir,
            "add" or "remove" => outputPath,
            _ => path
        };

        return ResultHelper.FinalizeResult((dynamic)result, effectiveOutputPath, (string?)null);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters.
    /// </summary>
    /// <param name="path">The input email file path.</param>
    /// <param name="outputPath">The output email file path.</param>
    /// <param name="attachmentPath">The attachment file path.</param>
    /// <param name="outputDir">The output directory for extraction.</param>
    /// <param name="index">The zero-based attachment index.</param>
    /// <returns>OperationParameters configured for the email attachment operation.</returns>
    private static OperationParameters BuildParameters(
        string path,
        string? outputPath,
        string? attachmentPath,
        string? outputDir,
        int? index)
    {
        var parameters = new OperationParameters();
        parameters.Set("path", path);
        parameters.SetIfNotNull("outputPath", outputPath);
        parameters.SetIfNotNull("attachmentPath", attachmentPath);
        parameters.SetIfNotNull("outputDir", outputDir);
        parameters.SetIfHasValue("index", index);
        return parameters;
    }
}
