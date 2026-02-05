using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Pdf.Compliance;

namespace AsposeMcpServer.Handlers.Pdf.Compliance;

/// <summary>
///     Handler for converting a PDF document to a specified compliance format such as PDF/A or PDF/UA.
/// </summary>
[ResultType(typeof(ConvertCompliancePdfResult))]
public class ConvertPdfComplianceHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "convert";

    /// <summary>
    ///     Converts a PDF document to the specified compliance format.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: format (target compliance format)
    ///     Optional: logPath (path to write conversion log)
    /// </param>
    /// <returns>Conversion result with success status.</returns>
    /// <exception cref="ArgumentException">Thrown when format is unsupported or logPath is invalid.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var format = parameters.GetRequired<string>("format");
        var logPath = parameters.GetOptional<string?>("logPath");

        var document = context.Document;
        var pdfFormat = ValidatePdfComplianceHandler.ResolvePdfFormat(format);

        if (logPath != null)
            SecurityHelper.ValidateFilePath(logPath, "logPath", true);

        var tempLog = logPath ?? Path.GetTempFileName();
        var success = document.Convert(tempLog, pdfFormat, ConvertErrorAction.Delete);

        if (logPath == null && File.Exists(tempLog))
            File.Delete(tempLog);

        MarkModified(context);

        var formatName = FormatToDisplayName(format);
        var message = success
            ? $"Document successfully converted to {formatName}."
            : $"Document conversion to {formatName} completed with errors.";

        return new ConvertCompliancePdfResult
        {
            Format = formatName,
            IsSuccess = success,
            LogPath = logPath,
            Message = message
        };
    }

    /// <summary>
    ///     Converts a format string to a human-readable display name.
    /// </summary>
    /// <param name="format">The format string.</param>
    /// <returns>The display name for the format.</returns>
    private static string FormatToDisplayName(string format)
    {
        return format.ToLowerInvariant().Replace(" ", "") switch
        {
            "pdf/a-1a" or "pdfa1a" => "PDF/A-1a",
            "pdf/a-1b" or "pdfa1b" => "PDF/A-1b",
            "pdf/a-2a" or "pdfa2a" => "PDF/A-2a",
            "pdf/a-2b" or "pdfa2b" => "PDF/A-2b",
            "pdf/a-3a" or "pdfa3a" => "PDF/A-3a",
            "pdf/a-3b" or "pdfa3b" => "PDF/A-3b",
            "pdf/ua-1" or "pdfua1" => "PDF/UA-1",
            _ => format
        };
    }
}
