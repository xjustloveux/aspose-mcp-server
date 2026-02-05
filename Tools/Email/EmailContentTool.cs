using System.ComponentModel;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Email;

/// <summary>
///     Tool for email content operations including body, headers, subject, and recipients management.
///     Email tools operate directly on files without using DocumentContext or session management.
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Email.Content")]
[McpServerToolType]
public class EmailContentTool
{
    /// <summary>
    ///     Handler registry for email content operations.
    /// </summary>
    private readonly HandlerRegistry<object> _handlerRegistry;

    /// <summary>
    ///     Initializes a new instance of the <see cref="EmailContentTool" /> class.
    /// </summary>
    public EmailContentTool()
    {
        _handlerRegistry =
            HandlerRegistry<object>.CreateFromNamespace("AsposeMcpServer.Handlers.Email.Content");
    }

    /// <summary>
    ///     Executes an email content operation (get_body, set_body, get_headers, set_headers, get_subject, set_subject,
    ///     get_recipients, set_recipients).
    /// </summary>
    /// <param name="operation">
    ///     The operation to perform: get_body, set_body, get_headers, set_headers, get_subject, set_subject,
    ///     get_recipients, set_recipients.
    /// </param>
    /// <param name="path">Input email file path (required for all operations).</param>
    /// <param name="outputPath">Output file path for save operations (defaults to path).</param>
    /// <param name="body">Body content (for set_body).</param>
    /// <param name="isHtml">Whether the body is HTML (for set_body, default: false).</param>
    /// <param name="subject">Subject text (for set_subject).</param>
    /// <param name="name">Header name (for set_headers).</param>
    /// <param name="value">Header value (for set_headers).</param>
    /// <param name="from">Sender address (for set_recipients).</param>
    /// <param name="to">Comma-separated To addresses (for set_recipients).</param>
    /// <param name="cc">Comma-separated CC addresses (for set_recipients).</param>
    /// <param name="bcc">Comma-separated BCC addresses (for set_recipients).</param>
    /// <returns>Operation result depending on the operation type.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "email_content",
        Title = "Email Content Operations",
        Destructive = false,
        Idempotent = true,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(
        @"Perform email content operations. Supports 8 operations: get_body, set_body, get_headers, set_headers, get_subject, set_subject, get_recipients, set_recipients.

Usage examples:
- Get body: email_content(operation='get_body', path='email.eml')
- Set body: email_content(operation='set_body', path='email.eml', body='New content')
- Set HTML body: email_content(operation='set_body', path='email.eml', body='<h1>Hello</h1>', isHtml=true)
- Get headers: email_content(operation='get_headers', path='email.eml')
- Set header: email_content(operation='set_headers', path='email.eml', name='X-Custom', value='test')
- Get subject: email_content(operation='get_subject', path='email.eml')
- Set subject: email_content(operation='set_subject', path='email.eml', subject='New Subject')
- Get recipients: email_content(operation='get_recipients', path='email.eml')
- Set recipients: email_content(operation='set_recipients', path='email.eml', to='a@b.com,c@d.com', cc='e@f.com')

Supported email formats: EML, MSG, MHTML/MHT, HTML")]
    public object Execute(
        [Description(@"Operation to perform.
- 'get_body': Get email body content (required params: path)
- 'set_body': Set email body content (required params: path, body; optional: outputPath, isHtml)
- 'get_headers': Get all email headers (required params: path)
- 'set_headers': Set an email header (required params: path, name, value; optional: outputPath)
- 'get_subject': Get email subject (required params: path)
- 'set_subject': Set email subject (required params: path, subject; optional: outputPath)
- 'get_recipients': Get email recipients (required params: path)
- 'set_recipients': Set email recipients (required params: path; optional: from, to, cc, bcc, outputPath)")]
        string operation,
        [Description("Input email file path (required for all operations)")]
        string path,
        [Description("Output file path for save operations (defaults to input path)")]
        string? outputPath = null,
        [Description("Body content (for set_body)")]
        string? body = null,
        [Description("Whether the body is HTML content (for set_body, default: false)")]
        bool isHtml = false,
        [Description("Subject text (for set_subject)")]
        string? subject = null,
        [Description("Header name (for set_headers)")]
        string? name = null,
        [Description("Header value (for set_headers)")]
        string? value = null,
        [Description("Sender email address (for set_recipients)")]
        string? from = null,
        [Description("Comma-separated To email addresses (for set_recipients)")]
        string? to = null,
        [Description("Comma-separated CC email addresses (for set_recipients)")]
        string? cc = null,
        [Description("Comma-separated BCC email addresses (for set_recipients)")]
        string? bcc = null)
    {
        var parameters = BuildParameters(path, outputPath, body, isHtml, subject, name, value, from, to, cc, bcc);
        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<object>
        {
            Document = new object(),
            SourcePath = path,
            OutputPath = outputPath
        };

        var result = handler.Execute(operationContext, parameters);
        var effectiveOutputPath = IsWriteOperation(operation) ? outputPath ?? path : path;

        return ResultHelper.FinalizeResult((dynamic)result, effectiveOutputPath, (string?)null);
    }

    /// <summary>
    ///     Determines if the operation is a write operation that modifies the file.
    /// </summary>
    /// <param name="operation">The operation name.</param>
    /// <returns>True if the operation modifies the email file.</returns>
    private static bool IsWriteOperation(string operation)
    {
        return operation.StartsWith("set_", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters.
    /// </summary>
    /// <param name="path">The input file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="body">The email body content.</param>
    /// <param name="isHtml">Whether the body is HTML.</param>
    /// <param name="subject">The email subject.</param>
    /// <param name="name">The header name.</param>
    /// <param name="value">The header value.</param>
    /// <param name="from">The sender email address.</param>
    /// <param name="to">The recipient To addresses.</param>
    /// <param name="cc">The recipient CC addresses.</param>
    /// <param name="bcc">The recipient BCC addresses.</param>
    /// <returns>OperationParameters configured for the email content operation.</returns>
    private static OperationParameters BuildParameters(
        string path,
        string? outputPath,
        string? body,
        bool isHtml,
        string? subject,
        string? name,
        string? value,
        string? from,
        string? to,
        string? cc,
        string? bcc)
    {
        var parameters = new OperationParameters();
        parameters.Set("path", path);
        parameters.SetIfNotNull("outputPath", outputPath);
        parameters.SetIfNotNull("body", body);
        parameters.Set("isHtml", isHtml);
        parameters.SetIfNotNull("subject", subject);
        parameters.SetIfNotNull("name", name);
        parameters.SetIfNotNull("value", value);
        parameters.SetIfNotNull("from", from);
        parameters.SetIfNotNull("to", to);
        parameters.SetIfNotNull("cc", cc);
        parameters.SetIfNotNull("bcc", bcc);
        return parameters;
    }
}
