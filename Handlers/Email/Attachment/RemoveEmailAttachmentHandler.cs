using Aspose.Email;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Email.Attachment;

/// <summary>
///     Handler for removing an attachment from an email message by index.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class RemoveEmailAttachmentHandler : OperationHandlerBase<object>
{
    /// <inheritdoc />
    public override string Operation => "remove";

    /// <summary>
    ///     Removes an attachment at the specified index from the email and saves the result.
    /// </summary>
    /// <param name="context">The operation context (not used for email operations).</param>
    /// <param name="parameters">
    ///     Required: path (email file path), outputPath (output file path), index (zero-based attachment index).
    /// </param>
    /// <returns>A <see cref="SuccessResult" /> confirming the attachment was removed.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the email file does not exist.</exception>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when the index is out of range.</exception>
    public override object Execute(OperationContext<object> context, OperationParameters parameters)
    {
        var path = parameters.GetRequired<string>("path");
        var outputPath = parameters.GetRequired<string>("outputPath");
        var idx = parameters.GetRequired<int>("index");

        SecurityHelper.ValidateFilePath(path, "path", true);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        if (!File.Exists(path))
            throw new FileNotFoundException($"Email file not found: {path}");

        var message = MailMessage.Load(path);

        if (idx < 0 || idx >= message.Attachments.Count)
            // ReSharper disable once NotResolvedInText
            throw new ArgumentOutOfRangeException(
                "index",
                $"Attachment index {idx} is out of range. Email has {message.Attachments.Count} attachment(s).");

        var removedName = message.Attachments[idx].Name;
        message.Attachments.RemoveAt(idx);
        message.Save(outputPath, EmailFormatHelper.DetermineEmailSaveFormat(outputPath));

        return new SuccessResult
        {
            Message = $"Attachment '{removedName}' (index {idx}) removed successfully. " +
                      $"Email now has {message.Attachments.Count} attachment(s)."
        };
    }
}
