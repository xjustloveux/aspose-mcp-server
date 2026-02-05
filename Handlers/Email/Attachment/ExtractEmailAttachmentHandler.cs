using Aspose.Email;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Email.Attachment;

/// <summary>
///     Handler for extracting a specific attachment from an email message by index.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class ExtractEmailAttachmentHandler : OperationHandlerBase<object>
{
    /// <inheritdoc />
    public override string Operation => "extract";

    /// <summary>
    ///     Extracts a specific attachment from the email and saves it to the output directory.
    /// </summary>
    /// <param name="context">The operation context (not used for email operations).</param>
    /// <param name="parameters">
    ///     Required: path (email file path), outputDir (output directory), index (zero-based attachment index).
    /// </param>
    /// <returns>A <see cref="SuccessResult" /> confirming the attachment was extracted.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the email file does not exist.</exception>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when the index is out of range.</exception>
    public override object Execute(OperationContext<object> context, OperationParameters parameters)
    {
        var path = parameters.GetRequired<string>("path");
        var outputDir = parameters.GetRequired<string>("outputDir");
        var idx = parameters.GetRequired<int>("index");

        SecurityHelper.ValidateFilePath(path, "path", true);
        SecurityHelper.ValidateFilePath(outputDir, "outputDir", true);

        if (!File.Exists(path))
            throw new FileNotFoundException($"Email file not found: {path}");

        var message = MailMessage.Load(path);

        if (idx < 0 || idx >= message.Attachments.Count)
            // CA2208, S3928, NotResolvedInText - 'index' is a dynamic parameter from the parameters dictionary, not a method parameter
            // ReSharper disable once NotResolvedInText
#pragma warning disable CA2208, S3928
            throw new ArgumentOutOfRangeException(
                "index",
                $"Attachment index {idx} is out of range. Email has {message.Attachments.Count} attachment(s).");
#pragma warning restore CA2208, S3928

        Directory.CreateDirectory(outputDir);

        var attachment = message.Attachments[idx];
        var fileName = SecurityHelper.SanitizeFileName(attachment.Name);
        var outputPath = Path.Combine(outputDir, fileName);
        attachment.Save(outputPath);

        return new SuccessResult
        {
            Message = $"Attachment '{attachment.Name}' extracted to '{outputPath}'."
        };
    }
}
