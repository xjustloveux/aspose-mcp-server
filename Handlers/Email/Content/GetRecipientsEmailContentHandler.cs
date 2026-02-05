using Aspose.Email;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Email.Content;

namespace AsposeMcpServer.Handlers.Email.Content;

/// <summary>
///     Handler for retrieving all recipients of an email.
/// </summary>
[ResultType(typeof(EmailRecipientsResult))]
public class GetRecipientsEmailContentHandler : OperationHandlerBase<object>
{
    /// <inheritdoc />
    public override string Operation => "get_recipients";

    /// <summary>
    ///     Loads an email file and returns all recipient information (From, To, CC, BCC).
    /// </summary>
    /// <param name="context">The operation context (not used for email operations).</param>
    /// <param name="parameters">
    ///     Required: path (email file path to read).
    /// </param>
    /// <returns>An <see cref="EmailRecipientsResult" /> containing all recipient details.</returns>
    /// <exception cref="ArgumentException">Thrown when path is missing or invalid.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the email file does not exist.</exception>
    public override object Execute(OperationContext<object> context, OperationParameters parameters)
    {
        var path = parameters.GetRequired<string>("path");
        SecurityHelper.ValidateFilePath(path, "path", true);

        if (!File.Exists(path))
            throw new FileNotFoundException($"Email file not found: {path}");

        var message = MailMessage.Load(path);

        var toList = message.To.Select(a => a.Address).ToList();
        var ccList = message.CC.Select(a => a.Address).ToList();
        var bccList = message.Bcc.Select(a => a.Address).ToList();
        var totalCount = toList.Count + ccList.Count + bccList.Count;

        return new EmailRecipientsResult
        {
            From = message.From?.Address,
            To = toList,
            Cc = ccList,
            Bcc = bccList,
            Message = $"Retrieved {totalCount} recipient(s) from email"
        };
    }
}
