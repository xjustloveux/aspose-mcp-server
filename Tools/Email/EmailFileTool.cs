using System.ComponentModel;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Email;

/// <summary>
///     Tool for email file operations including create, load, save, convert, and format detection.
///     Email tools operate directly on files without using DocumentContext or session management.
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Email.FileOperations")]
[McpServerToolType]
public class EmailFileTool
{
    /// <summary>
    ///     Handler registry for email file operations.
    /// </summary>
    private readonly HandlerRegistry<object> _handlerRegistry;

    /// <summary>
    ///     Initializes a new instance of the <see cref="EmailFileTool" /> class.
    /// </summary>
    public EmailFileTool()
    {
        _handlerRegistry =
            HandlerRegistry<object>.CreateFromNamespace("AsposeMcpServer.Handlers.Email.FileOperations");
    }

    /// <summary>
    ///     Executes an email file operation (create, get_info, save, convert, detect_format).
    /// </summary>
    /// <param name="operation">
    ///     The operation to perform: create, get_info, save, convert, detect_format.
    /// </param>
    /// <param name="path">Input email file path (required for get_info, save, convert, detect_format).</param>
    /// <param name="outputPath">Output file path (required for create, save, convert).</param>
    /// <param name="subject">Email subject (for create).</param>
    /// <param name="body">Email body content (for create).</param>
    /// <param name="from">Sender email address (for create).</param>
    /// <param name="to">Recipient email address (for create).</param>
    /// <param name="isHtml">Whether the body is HTML content (for create, default: false).</param>
    /// <returns>Operation result depending on the operation type.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "email_file",
        Title = "Email File Operations",
        Destructive = false,
        Idempotent = true,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(@"Perform email file operations. Supports 5 operations: create, get_info, save, convert, detect_format.

Usage examples:
- Create email: email_file(operation='create', outputPath='email.eml', subject='Hello', body='World', from='a@b.com', to='c@d.com')
- Get email info: email_file(operation='get_info', path='email.eml')
- Save email: email_file(operation='save', path='email.eml', outputPath='copy.eml')
- Convert email: email_file(operation='convert', path='email.eml', outputPath='email.msg')
- Detect format: email_file(operation='detect_format', path='email.eml')

Supported email formats: EML, MSG, MHTML/MHT, HTML")]
    public object Execute(
        [Description(@"Operation to perform.
- 'create': Create a new email file (required params: outputPath; optional: subject, body, from, to, isHtml)
- 'get_info': Load email and return metadata (required params: path)
- 'save': Save email to a new location (required params: path, outputPath)
- 'convert': Convert email to another format (required params: path, outputPath)
- 'detect_format': Detect the format of an email file (required params: path)")]
        string operation,
        [Description("Input email file path (required for get_info, save, convert, detect_format)")]
        string? path = null,
        [Description("Output file path (required for create, save, convert)")]
        string? outputPath = null,
        [Description("Email subject (for create)")]
        string? subject = null,
        [Description("Email body content (for create)")]
        string? body = null,
        [Description("Sender email address (for create)")]
        string? from = null,
        [Description("Recipient email address (for create)")]
        string? to = null,
        [Description("Whether the body is HTML content (for create, default: false)")]
        bool isHtml = false)
    {
        var parameters = BuildParameters(path, outputPath, subject, body, from, to, isHtml);
        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<object>
        {
            Document = new object(),
            SourcePath = path,
            OutputPath = outputPath
        };

        var result = handler.Execute(operationContext, parameters);
        var effectiveOutputPath = ResolveOutputPath(operation, path, outputPath);

        return ResultHelper.FinalizeResult((dynamic)result, effectiveOutputPath, (string?)null);
    }

    /// <summary>
    ///     Resolves the effective output path based on the operation type.
    /// </summary>
    /// <param name="operation">The operation name.</param>
    /// <param name="path">The input file path.</param>
    /// <param name="outputPath">The explicit output file path.</param>
    /// <returns>The effective output path for the result.</returns>
    private static string? ResolveOutputPath(string operation, string? path, string? outputPath)
    {
        return operation.ToLowerInvariant() switch
        {
            "create" => outputPath,
            "save" or "convert" => outputPath,
            _ => path
        };
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters.
    /// </summary>
    /// <param name="path">The input file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="subject">The email subject.</param>
    /// <param name="body">The email body content.</param>
    /// <param name="from">The sender email address.</param>
    /// <param name="to">The recipient email address.</param>
    /// <param name="isHtml">Whether the body is HTML.</param>
    /// <returns>OperationParameters configured for the email file operation.</returns>
    private static OperationParameters BuildParameters(
        string? path,
        string? outputPath,
        string? subject,
        string? body,
        string? from,
        string? to,
        bool isHtml)
    {
        var parameters = new OperationParameters();
        parameters.SetIfNotNull("path", path);
        parameters.SetIfNotNull("outputPath", outputPath);
        parameters.SetIfNotNull("subject", subject);
        parameters.SetIfNotNull("body", body);
        parameters.SetIfNotNull("from", from);
        parameters.SetIfNotNull("to", to);
        parameters.Set("isHtml", isHtml);
        return parameters;
    }
}
