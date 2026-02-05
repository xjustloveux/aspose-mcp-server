using Aspose.Email;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Email.FileOperations;

namespace AsposeMcpServer.Handlers.Email.FileOperations;

/// <summary>
///     Handler for loading an email file and returning its metadata.
/// </summary>
[ResultType(typeof(EmailFileInfo))]
public class LoadEmailFileHandler : OperationHandlerBase<object>
{
    /// <inheritdoc />
    public override string Operation => "get_info";

    /// <summary>
    ///     Loads an email file and returns its metadata information.
    /// </summary>
    /// <param name="context">The operation context (not used for email operations).</param>
    /// <param name="parameters">
    ///     Required: path (email file path to load).
    /// </param>
    /// <returns>An <see cref="EmailFileInfo" /> containing the email metadata.</returns>
    /// <exception cref="ArgumentException">Thrown when path is missing or invalid.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the email file does not exist.</exception>
    public override object Execute(OperationContext<object> context, OperationParameters parameters)
    {
        var path = parameters.GetRequired<string>("path");
        SecurityHelper.ValidateFilePath(path, "path", true);

        if (!File.Exists(path))
            throw new FileNotFoundException($"Email file not found: {path}");

        var message = MailMessage.Load(path);

        var ext = Path.GetExtension(path).ToLowerInvariant();
        var format = ext switch
        {
            ".eml" => "EML",
            ".msg" => "MSG",
            ".mhtml" or ".mht" => "MHTML",
            ".html" or ".htm" => "HTML",
            _ => "Unknown"
        };

        return new EmailFileInfo
        {
            Subject = message.Subject,
            From = message.From?.Address,
            To = message.To.Count > 0 ? string.Join(", ", message.To.Select(a => a.Address)) : null,
            Date = message.Date != DateTime.MinValue ? message.Date.ToString("O") : null,
            Format = format,
            HasAttachments = message.Attachments.Count > 0,
            AttachmentCount = message.Attachments.Count,
            Message = $"Email loaded from {path}"
        };
    }
}
