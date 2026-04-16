using Aspose.Email;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Email.Attachment;

/// <summary>
///     Handler for adding an attachment to an email message.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class AddEmailAttachmentHandler : OperationHandlerBase<object>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds an attachment to the specified email file and saves the result.
    /// </summary>
    /// <param name="context">The operation context (not used for email operations).</param>
    /// <param name="parameters">
    ///     Required: path (email file path), outputPath (output file path), attachmentPath (file to attach).
    /// </param>
    /// <returns>A <see cref="SuccessResult" /> confirming the attachment was added.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the email file or attachment file does not exist.</exception>
    public override object Execute(OperationContext<object> context, OperationParameters parameters)
    {
        var path = parameters.GetRequired<string>("path");
        var outputPath = parameters.GetRequired<string>("outputPath");
        var attachmentPath = parameters.GetRequired<string>("attachmentPath");

        SecurityHelper.ValidateFilePath(path, "path", true);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);
        SecurityHelper.ValidateFilePath(attachmentPath, "attachmentPath", true);

        // Allowlist + symlink resolution for read sinks (bug 20260416-handler-allowlist-bypass).
        var allowedBases = context.ServerConfig?.AllowedBasePaths ?? [];
        path = SecurityHelper.ResolveAndEnsureWithinAllowlist(path, allowedBases, nameof(path));
        attachmentPath =
            SecurityHelper.ResolveAndEnsureWithinAllowlist(attachmentPath, allowedBases, nameof(attachmentPath));

        if (!File.Exists(path))
            throw new FileNotFoundException("The specified file was not found.");

        if (!File.Exists(attachmentPath))
            throw new FileNotFoundException("The specified file was not found.");

        var message = MailMessage.Load(path);
        var attachment = new Aspose.Email.Attachment(attachmentPath);
        message.Attachments.Add(attachment);
        // H41: resolve symlinks immediately before the sink (bug 20260415-symlink-toctou-sweep).
        outputPath = SecurityHelper.ResolveAndEnsureWithinAllowlist(outputPath,
            context.ServerConfig?.AllowedBasePaths ?? [], nameof(outputPath));
        message.Save(outputPath, EmailFormatHelper.DetermineEmailSaveFormat(outputPath));

        return new SuccessResult
        {
            Message = $"Attachment '{Path.GetFileName(attachmentPath)}' added successfully. " +
                      $"Email now has {message.Attachments.Count} attachment(s)."
        };
    }
}
