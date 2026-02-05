using Aspose.Email;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Handlers.Email.FileOperations;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Email.Content;

/// <summary>
///     Handler for setting the body content of an email.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class SetBodyEmailContentHandler : OperationHandlerBase<object>
{
    /// <inheritdoc />
    public override string Operation => "set_body";

    /// <summary>
    ///     Loads an email file, sets its body content, and saves it.
    /// </summary>
    /// <param name="context">The operation context (not used for email operations).</param>
    /// <param name="parameters">
    ///     Required: path (email file path), body (new body content).
    ///     Optional: outputPath (save location, defaults to path), isHtml (default: false).
    /// </param>
    /// <returns>A <see cref="SuccessResult" /> indicating the body was set successfully.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the email file does not exist.</exception>
    public override object Execute(OperationContext<object> context, OperationParameters parameters)
    {
        var path = parameters.GetRequired<string>("path");
        var body = parameters.GetRequired<string>("body");
        var outputPath = parameters.GetOptional("outputPath", path);
        var isHtml = parameters.GetOptional("isHtml", false);
        SecurityHelper.ValidateFilePath(path, "path", true);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        if (!File.Exists(path))
            throw new FileNotFoundException($"Email file not found: {path}");

        var message = MailMessage.Load(path);

        if (isHtml)
            message.HtmlBody = body;
        else
            message.Body = body;

        var saveOptions = CreateEmailFileHandler.DetectSaveOptions(outputPath);
        message.Save(outputPath, saveOptions);

        return new SuccessResult
        {
            Message = isHtml
                ? $"Email HTML body updated and saved to {outputPath}"
                : $"Email body updated and saved to {outputPath}"
        };
    }
}
