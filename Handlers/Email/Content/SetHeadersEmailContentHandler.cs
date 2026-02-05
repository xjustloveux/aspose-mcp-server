using Aspose.Email;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Handlers.Email.FileOperations;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Email.Content;

/// <summary>
///     Handler for setting an email header value.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class SetHeadersEmailContentHandler : OperationHandlerBase<object>
{
    /// <inheritdoc />
    public override string Operation => "set_headers";

    /// <summary>
    ///     Loads an email file, sets or updates a header value, and saves it.
    /// </summary>
    /// <param name="context">The operation context (not used for email operations).</param>
    /// <param name="parameters">
    ///     Required: path (email file path), name (header name), value (header value).
    ///     Optional: outputPath (save location, defaults to path).
    /// </param>
    /// <returns>A <see cref="SuccessResult" /> indicating the header was set successfully.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the email file does not exist.</exception>
    public override object Execute(OperationContext<object> context, OperationParameters parameters)
    {
        var path = parameters.GetRequired<string>("path");
        var name = parameters.GetRequired<string>("name");
        var value = parameters.GetRequired<string>("value");
        var outputPath = parameters.GetOptional("outputPath", path);
        SecurityHelper.ValidateFilePath(path, "path", true);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        if (!File.Exists(path))
            throw new FileNotFoundException($"Email file not found: {path}");

        var message = MailMessage.Load(path);
        message.Headers.Set(name, value);

        var saveOptions = CreateEmailFileHandler.DetectSaveOptions(outputPath);
        message.Save(outputPath, saveOptions);

        return new SuccessResult
        {
            Message = $"Header '{name}' set to '{value}' and saved to {outputPath}"
        };
    }
}
