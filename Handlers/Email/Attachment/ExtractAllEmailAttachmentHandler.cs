using Aspose.Email;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Email.Attachment;

/// <summary>
///     Handler for extracting all attachments from an email message.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class ExtractAllEmailAttachmentHandler : OperationHandlerBase<object>
{
    /// <inheritdoc />
    public override string Operation => "extract_all";

    /// <summary>
    ///     Extracts all attachments from the email and saves them to the output directory.
    /// </summary>
    /// <param name="context">The operation context (not used for email operations).</param>
    /// <param name="parameters">
    ///     Required: path (email file path), outputDir (output directory).
    /// </param>
    /// <returns>A <see cref="SuccessResult" /> confirming all attachments were extracted.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the email file does not exist.</exception>
    public override object Execute(OperationContext<object> context, OperationParameters parameters)
    {
        var path = parameters.GetRequired<string>("path");
        var outputDir = parameters.GetRequired<string>("outputDir");

        SecurityHelper.ValidateFilePath(path, "path", true);
        SecurityHelper.ValidateFilePath(outputDir, "outputDir", true);

        if (!File.Exists(path))
            throw new FileNotFoundException($"Email file not found: {path}");

        var message = MailMessage.Load(path);

        if (message.Attachments.Count == 0)
            return new SuccessResult
            {
                Message = "No attachments found in the email."
            };

        Directory.CreateDirectory(outputDir);

        var extractedFiles = new List<string>();
        foreach (var attachment in message.Attachments)
        {
            var fileName = SecurityHelper.SanitizeFileName(attachment.Name);
            var outputPath = Path.Combine(outputDir, fileName);
            attachment.Save(outputPath);
            extractedFiles.Add(fileName);
        }

        return new SuccessResult
        {
            Message = $"Extracted {extractedFiles.Count} attachment(s) to '{outputDir}': " +
                      string.Join(", ", extractedFiles) + "."
        };
    }
}
