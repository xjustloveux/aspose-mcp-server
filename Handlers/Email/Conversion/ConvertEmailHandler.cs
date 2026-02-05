using Aspose.Email;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Email.Conversion;

namespace AsposeMcpServer.Handlers.Email.Conversion;

/// <summary>
///     Handler for converting email messages between different formats.
/// </summary>
[ResultType(typeof(EmailConversionResult))]
public class ConvertEmailHandler : OperationHandlerBase<object>
{
    /// <inheritdoc />
    public override string Operation => "convert";

    /// <summary>
    ///     Converts an email message from one format to another.
    /// </summary>
    /// <param name="context">The operation context (not used for email operations).</param>
    /// <param name="parameters">
    ///     Required: path (source email file), outputPath (destination file).
    /// </param>
    /// <returns>An <see cref="EmailConversionResult" /> containing conversion details.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing, invalid, or format is unsupported.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the input file does not exist.</exception>
    public override object Execute(OperationContext<object> context, OperationParameters parameters)
    {
        var path = parameters.GetRequired<string>("path");
        var outputPath = parameters.GetRequired<string>("outputPath");

        SecurityHelper.ValidateFilePath(path, "path", true);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        if (!File.Exists(path))
            throw new FileNotFoundException($"Input file not found: {path}");

        var message = MailMessage.Load(path);

        var outputExt = Path.GetExtension(outputPath).ToLowerInvariant();
        var saveOptions = GetSaveOptions(outputExt);

        message.Save(outputPath, saveOptions);

        var inputExt = Path.GetExtension(path).ToLowerInvariant();

        return new EmailConversionResult
        {
            SourcePath = path,
            OutputPath = outputPath,
            SourceFormat = inputExt.TrimStart('.').ToUpperInvariant(),
            TargetFormat = outputExt.TrimStart('.').ToUpperInvariant(),
            FileSize = File.Exists(outputPath) ? new FileInfo(outputPath).Length : null,
            Message = $"Email converted from {inputExt} to {outputExt}"
        };
    }

    /// <summary>
    ///     Gets the appropriate save options for the specified output file extension.
    /// </summary>
    /// <param name="extension">The output file extension (including the dot).</param>
    /// <returns>The <see cref="SaveOptions" /> for saving the email in the target format.</returns>
    /// <exception cref="ArgumentException">Thrown when the target format is not supported.</exception>
    private static SaveOptions GetSaveOptions(string extension)
    {
        return extension switch
        {
            ".eml" => SaveOptions.DefaultEml,
            ".emlx" => new EmlSaveOptions(MailMessageSaveType.EmlxFormat),
            ".msg" => SaveOptions.DefaultMsgUnicode,
            ".mht" or ".mhtml" => SaveOptions.DefaultMhtml,
            ".html" or ".htm" => SaveOptions.DefaultHtml,
            _ => throw new ArgumentException(
                $"Unsupported target format: {extension}. Supported: eml, emlx, msg, mht, mhtml, html, htm")
        };
    }
}
