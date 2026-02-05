using Aspose.Email;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Handlers.Email.FileOperations;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Email.Content;

/// <summary>
///     Handler for setting the subject of an email.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class SetSubjectEmailContentHandler : OperationHandlerBase<object>
{
    /// <inheritdoc />
    public override string Operation => "set_subject";

    /// <summary>
    ///     Loads an email file, sets its subject, and saves it.
    /// </summary>
    /// <param name="context">The operation context (not used for email operations).</param>
    /// <param name="parameters">
    ///     Required: path (email file path), subject (new subject text).
    ///     Optional: outputPath (save location, defaults to path).
    /// </param>
    /// <returns>A <see cref="SuccessResult" /> indicating the subject was set successfully.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the email file does not exist.</exception>
    public override object Execute(OperationContext<object> context, OperationParameters parameters)
    {
        var path = parameters.GetRequired<string>("path");
        var subject = parameters.GetRequired<string>("subject");
        var outputPath = parameters.GetOptional("outputPath", path);
        SecurityHelper.ValidateFilePath(path, "path", true);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        if (!File.Exists(path))
            throw new FileNotFoundException($"Email file not found: {path}");

        var message = MailMessage.Load(path);
        message.Subject = subject;

        var saveOptions = CreateEmailFileHandler.DetectSaveOptions(outputPath);
        message.Save(outputPath, saveOptions);

        return new SuccessResult
        {
            Message = $"Email subject set to '{subject}' and saved to {outputPath}"
        };
    }
}
