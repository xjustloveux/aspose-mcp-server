using Aspose.Email;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Email.FileOperations;

namespace AsposeMcpServer.Handlers.Email.FileOperations;

/// <summary>
///     Handler for converting an email file from one format to another.
/// </summary>
[ResultType(typeof(EmailConversionResult))]
public class ConvertEmailFileHandler : OperationHandlerBase<object>
{
    /// <inheritdoc />
    public override string Operation => "convert";

    /// <summary>
    ///     Converts an email file to a different format based on the output file extension.
    /// </summary>
    /// <param name="context">The operation context (not used for email operations).</param>
    /// <param name="parameters">
    ///     Required: path (source email file path), outputPath (destination file path with target extension).
    /// </param>
    /// <returns>An <see cref="EmailConversionResult" /> containing conversion details.</returns>
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

        var sourceFormat = DetectFormatFromExtension(path);
        var targetFormat = DetectFormatFromExtension(outputPath);

        var message = MailMessage.Load(path);
        var saveOptions = CreateEmailFileHandler.DetectSaveOptions(outputPath);
        message.Save(outputPath, saveOptions);

        return new EmailConversionResult
        {
            SourcePath = path,
            OutputPath = outputPath,
            SourceFormat = sourceFormat,
            TargetFormat = targetFormat,
            Message = $"Email converted from {sourceFormat} to {targetFormat}"
        };
    }

    /// <summary>
    ///     Detects the email format name from the file extension.
    /// </summary>
    /// <param name="path">The file path.</param>
    /// <returns>The format name string.</returns>
    private static string DetectFormatFromExtension(string path)
    {
        var ext = Path.GetExtension(path).ToLowerInvariant();
        return ext switch
        {
            ".eml" => "EML",
            ".msg" => "MSG",
            ".mhtml" or ".mht" => "MHTML",
            ".html" or ".htm" => "HTML",
            _ => "EML"
        };
    }
}
