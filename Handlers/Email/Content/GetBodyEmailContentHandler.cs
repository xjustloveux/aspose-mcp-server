using Aspose.Email;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Email.Content;

namespace AsposeMcpServer.Handlers.Email.Content;

/// <summary>
///     Handler for retrieving the body content of an email.
/// </summary>
[ResultType(typeof(EmailBodyResult))]
public class GetBodyEmailContentHandler : OperationHandlerBase<object>
{
    /// <inheritdoc />
    public override string Operation => "get_body";

    /// <summary>
    ///     Loads an email file and returns its body content.
    /// </summary>
    /// <param name="context">The operation context (not used for email operations).</param>
    /// <param name="parameters">
    ///     Required: path (email file path to read).
    /// </param>
    /// <returns>An <see cref="EmailBodyResult" /> containing the email body.</returns>
    /// <exception cref="ArgumentException">Thrown when path is missing or invalid.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the email file does not exist.</exception>
    public override object Execute(OperationContext<object> context, OperationParameters parameters)
    {
        var path = parameters.GetRequired<string>("path");
        SecurityHelper.ValidateFilePath(path, "path", true);

        if (!File.Exists(path))
            throw new FileNotFoundException($"Email file not found: {path}");

        var message = MailMessage.Load(path);
        var isHtml = !string.IsNullOrEmpty(message.HtmlBody);

        return new EmailBodyResult
        {
            Body = message.Body,
            HtmlBody = message.HtmlBody,
            IsHtml = isHtml,
            Message = isHtml ? "Email body retrieved (HTML)" : "Email body retrieved (plain text)"
        };
    }
}
