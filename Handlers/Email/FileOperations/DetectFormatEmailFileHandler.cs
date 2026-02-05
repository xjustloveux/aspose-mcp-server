using Aspose.Email;
using Aspose.Email.Tools;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Email.FileOperations;

namespace AsposeMcpServer.Handlers.Email.FileOperations;

/// <summary>
///     Handler for detecting the format of an email file.
/// </summary>
[ResultType(typeof(DetectFormatEmailResult))]
public class DetectFormatEmailFileHandler : OperationHandlerBase<object>
{
    /// <inheritdoc />
    public override string Operation => "detect_format";

    /// <summary>
    ///     Detects the format of an email file using Aspose.Email's format detection.
    /// </summary>
    /// <param name="context">The operation context (not used for email operations).</param>
    /// <param name="parameters">
    ///     Required: path (email file path to analyze).
    /// </param>
    /// <returns>A <see cref="DetectFormatEmailResult" /> containing the detected format information.</returns>
    /// <exception cref="ArgumentException">Thrown when path is missing or invalid.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the email file does not exist.</exception>
    public override object Execute(OperationContext<object> context, OperationParameters parameters)
    {
        var path = parameters.GetRequired<string>("path");
        SecurityHelper.ValidateFilePath(path, "path", true);

        if (!File.Exists(path))
            throw new FileNotFoundException($"Email file not found: {path}");

        var formatInfo = FileFormatUtil.DetectFileFormat(path);
        var formatName = formatInfo.FileFormatType.ToString();
        var extension = MapFormatToExtension(formatInfo.FileFormatType);

        return new DetectFormatEmailResult
        {
            Format = formatName,
            Extension = extension,
            Message = $"Detected format: {formatName} (extension: {extension})"
        };
    }

    /// <summary>
    ///     Maps a <see cref="FileFormatType" /> to its typical file extension.
    /// </summary>
    /// <param name="formatType">The detected file format type.</param>
    /// <returns>The file extension string including the leading dot.</returns>
    private static string MapFormatToExtension(FileFormatType formatType)
    {
        return formatType switch
        {
            FileFormatType.Eml => ".eml",
            FileFormatType.Msg => ".msg",
            FileFormatType.Mht => ".mht",
            FileFormatType.Emlx => ".emlx",
            FileFormatType.Ics => ".ics",
            FileFormatType.Vcf => ".vcf",
            FileFormatType.Ost => ".ost",
            FileFormatType.Pst => ".pst",
            FileFormatType.Mbox => ".mbox",
            FileFormatType.Oft => ".oft",
            FileFormatType.Olm => ".olm",
            FileFormatType.Tnef => ".tnef",
            _ => ".unknown"
        };
    }
}
