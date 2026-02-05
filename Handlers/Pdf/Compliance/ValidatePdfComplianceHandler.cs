using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Pdf.Compliance;

namespace AsposeMcpServer.Handlers.Pdf.Compliance;

/// <summary>
///     Handler for validating PDF document compliance against standards such as PDF/A and PDF/UA.
/// </summary>
[ResultType(typeof(ValidateCompliancePdfResult))]
public class ValidatePdfComplianceHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "validate";

    /// <summary>
    ///     Validates a PDF document against the specified compliance format.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: format (compliance format to validate against)
    ///     Optional: logPath (path to write validation log)
    /// </param>
    /// <returns>Validation result with compliance status and error count.</returns>
    /// <exception cref="ArgumentException">Thrown when format is unsupported or logPath is invalid.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var format = parameters.GetRequired<string>("format");
        var logPath = parameters.GetOptional<string?>("logPath");

        var document = context.Document;

        var pdfFormat = ResolvePdfFormat(format);

        if (logPath != null)
            SecurityHelper.ValidateFilePath(logPath, "logPath", true);

        var tempLog = logPath ?? Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
        var isCompliant = document.Validate(tempLog, pdfFormat);

        var errorCount = 0;
        if (File.Exists(tempLog))
        {
            var logContent = File.ReadAllText(tempLog);
            if (!string.IsNullOrWhiteSpace(logContent))
                errorCount = logContent.Split('\n', StringSplitOptions.RemoveEmptyEntries).Length;
        }

        if (logPath == null && File.Exists(tempLog))
            File.Delete(tempLog);

        var formatName = FormatToDisplayName(format);
        var message = isCompliant
            ? $"Document is compliant with {formatName}."
            : $"Document is not compliant with {formatName}. Found {errorCount} error(s).";

        return new ValidateCompliancePdfResult
        {
            IsCompliant = isCompliant,
            Format = formatName,
            ErrorCount = errorCount,
            LogPath = logPath,
            Message = message
        };
    }

    /// <summary>
    ///     Resolves a string format name to the corresponding <see cref="PdfFormat" /> enum value.
    /// </summary>
    /// <param name="format">The format string (e.g., "pdf/a-1b", "pdfa1b", "pdf/ua-1").</param>
    /// <returns>The resolved <see cref="PdfFormat" /> enum value.</returns>
    /// <exception cref="ArgumentException">Thrown when the format string is not recognized.</exception>
    internal static PdfFormat ResolvePdfFormat(string format)
    {
        return format.ToLowerInvariant().Replace(" ", "") switch
        {
            "pdf/a-1a" or "pdfa1a" => PdfFormat.PDF_A_1A,
            "pdf/a-1b" or "pdfa1b" => PdfFormat.PDF_A_1B,
            "pdf/a-2a" or "pdfa2a" => PdfFormat.PDF_A_2A,
            "pdf/a-2b" or "pdfa2b" => PdfFormat.PDF_A_2B,
            "pdf/a-3a" or "pdfa3a" => PdfFormat.PDF_A_3A,
            "pdf/a-3b" or "pdfa3b" => PdfFormat.PDF_A_3B,
            "pdf/ua-1" or "pdfua1" => PdfFormat.PDF_UA_1,
            _ => throw new ArgumentException(
                $"Unsupported compliance format: '{format}'. Supported formats: pdf/a-1a, pdf/a-1b, pdf/a-2a, pdf/a-2b, pdf/a-3a, pdf/a-3b, pdf/ua-1")
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
