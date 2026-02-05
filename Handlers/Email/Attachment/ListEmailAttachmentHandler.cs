using Aspose.Email;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Email.Attachment;

namespace AsposeMcpServer.Handlers.Email.Attachment;

/// <summary>
///     Handler for listing all attachments in an email message.
/// </summary>
[ResultType(typeof(GetAttachmentsEmailResult))]
public class ListEmailAttachmentHandler : OperationHandlerBase<object>
{
    /// <inheritdoc />
    public override string Operation => "list";

    /// <summary>
    ///     Lists all attachments in the specified email file.
    /// </summary>
    /// <param name="context">The operation context (not used for email operations).</param>
    /// <param name="parameters">
    ///     Required: path (email file path).
    /// </param>
    /// <returns>A <see cref="GetAttachmentsEmailResult" /> containing attachment information.</returns>
    /// <exception cref="ArgumentException">Thrown when the path parameter is missing or invalid.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the email file does not exist.</exception>
    public override object Execute(OperationContext<object> context, OperationParameters parameters)
    {
        var path = parameters.GetRequired<string>("path");
        SecurityHelper.ValidateFilePath(path, "path", true);

        if (!File.Exists(path))
            throw new FileNotFoundException($"Email file not found: {path}");

        var message = MailMessage.Load(path);

        var attachments = new List<AttachmentEmailInfo>();
        for (var i = 0; i < message.Attachments.Count; i++)
        {
            var att = message.Attachments[i];
            attachments.Add(new AttachmentEmailInfo
            {
                Index = i,
                Name = att.Name,
                ContentType = att.ContentType?.MediaType,
                Size = att.ContentStream?.Length ?? 0
            });
        }

        return new GetAttachmentsEmailResult
        {
            Count = attachments.Count,
            Attachments = attachments,
            Message = $"Found {attachments.Count} attachment(s) in the email."
        };
    }
}
