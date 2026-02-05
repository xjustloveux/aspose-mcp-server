using Aspose.Email;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Email.FileOperations;

/// <summary>
///     Handler for saving an email file to a specified path and format.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class SaveEmailFileHandler : OperationHandlerBase<object>
{
    /// <inheritdoc />
    public override string Operation => "save";

    /// <summary>
    ///     Loads an email from the source path and saves it to the output path.
    /// </summary>
    /// <param name="context">The operation context (not used for email operations).</param>
    /// <param name="parameters">
    ///     Required: path (source email file path), outputPath (destination file path).
    ///     Optional: format (target format, auto-detected from outputPath extension if not specified).
    /// </param>
    /// <returns>A <see cref="SuccessResult" /> indicating the email was saved successfully.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the source email file does not exist.</exception>
    public override object Execute(OperationContext<object> context, OperationParameters parameters)
    {
        var path = parameters.GetRequired<string>("path");
        var outputPath = parameters.GetRequired<string>("outputPath");
        SecurityHelper.ValidateFilePath(path, "path", true);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        if (!File.Exists(path))
            throw new FileNotFoundException($"Email file not found: {path}");

        var message = MailMessage.Load(path);
        var saveOptions = CreateEmailFileHandler.DetectSaveOptions(outputPath);
        message.Save(outputPath, saveOptions);

        return new SuccessResult
        {
            Message = $"Email saved to {outputPath}"
        };
    }
}
