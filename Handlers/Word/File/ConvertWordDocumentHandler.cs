using Aspose.Words;
using Aspose.Words.Saving;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Progress;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;
using ModelContextProtocol;

namespace AsposeMcpServer.Handlers.Word.File;

/// <summary>
///     Handler for converting Word documents to other formats.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class ConvertWordDocumentHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "convert";

    /// <summary>
    ///     Converts a Word document to another format.
    /// </summary>
    /// <param name="context">The operation context.</param>
    /// <param name="parameters">
    ///     Required: outputPath, either path or sessionId
    ///     Optional: format (inferred from outputPath if not provided)
    /// </param>
    /// <returns>Success message with conversion details.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or format is unsupported.</exception>
    /// <exception cref="InvalidOperationException">Thrown when session management is not enabled.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractConvertParameters(parameters);

        if (string.IsNullOrEmpty(p.Path) && string.IsNullOrEmpty(p.SessionId))
            throw new ArgumentException("Either path or sessionId is required for convert operation");
        if (string.IsNullOrEmpty(p.OutputPath))
            throw new ArgumentException("outputPath is required for convert operation");

        SecurityHelper.ValidateFilePath(p.OutputPath, "outputPath", true);

        var outputDir = Path.GetDirectoryName(p.OutputPath);
        if (!string.IsNullOrEmpty(outputDir))
            Directory.CreateDirectory(outputDir);

        var formatLower = p.Format?.ToLower();
        if (string.IsNullOrEmpty(formatLower))
        {
            var extension = Path.GetExtension(p.OutputPath).TrimStart('.').ToLower();
            formatLower = extension switch
            {
                "pdf" => "pdf",
                "html" or "htm" => "html",
                "docx" => "docx",
                "doc" => "doc",
                "txt" => "txt",
                "rtf" => "rtf",
                "odt" => "odt",
                "epub" => "epub",
                "xps" => "xps",
                _ => throw new ArgumentException(
                    $"Cannot infer format from extension '.{extension}'. Please specify format parameter.")
            };
        }

        Document doc;
        string sourceDescription;

        if (!string.IsNullOrEmpty(p.SessionId))
        {
            if (context.SessionManager == null)
                throw new InvalidOperationException("Session management is not enabled");

            var identity = context.IdentityAccessor?.GetCurrentIdentity() ?? SessionIdentity.GetAnonymous();
            doc = context.SessionManager.GetDocument<Document>(p.SessionId, identity);
            sourceDescription = $"session {p.SessionId}";
        }
        else
        {
            SecurityHelper.ValidateFilePath(p.Path!, allowAbsolutePaths: true);
            doc = new Document(p.Path);
            sourceDescription = p.Path!;
        }

        var saveFormat = formatLower switch
        {
            "pdf" => SaveFormat.Pdf,
            "html" => SaveFormat.Html,
            "docx" => SaveFormat.Docx,
            "doc" => SaveFormat.Doc,
            "txt" => SaveFormat.Text,
            "rtf" => SaveFormat.Rtf,
            "odt" => SaveFormat.Odt,
            "epub" => SaveFormat.Epub,
            "xps" => SaveFormat.Xps,
            _ => throw new ArgumentException($"Unsupported format: {p.Format}")
        };

        if (formatLower == "pdf")
        {
            var pdfSaveOptions = BuildPdfSaveOptions(p, context.Progress);
            doc.Save(p.OutputPath, pdfSaveOptions);
        }
        else
        {
            doc.Save(p.OutputPath, saveFormat);
        }

        return new SuccessResult
            { Message = $"Document converted from {sourceDescription} to {p.OutputPath} ({formatLower})" };
    }

    /// <summary>
    ///     Builds PdfSaveOptions with advanced settings from the convert parameters.
    /// </summary>
    /// <param name="p">The convert parameters.</param>
    /// <param name="progress">Optional progress reporter.</param>
    /// <returns>Configured PdfSaveOptions instance.</returns>
    /// <exception cref="ArgumentException">Thrown when the pdfCompliance value is unknown.</exception>
    private static PdfSaveOptions BuildPdfSaveOptions(ConvertParameters p,
        IProgress<ProgressNotificationValue>? progress)
    {
        var options = new PdfSaveOptions();

        if (progress != null)
            options.ProgressCallback = new WordsProgressAdapter(progress);

        if (!string.IsNullOrEmpty(p.PdfCompliance))
            options.Compliance = p.PdfCompliance.ToLowerInvariant() switch
            {
                "pdf17" => PdfCompliance.Pdf17,
                "pdfa1a" => PdfCompliance.PdfA1a,
                "pdfa1b" => PdfCompliance.PdfA1b,
                "pdfa2a" => PdfCompliance.PdfA2a,
                "pdfa2u" => PdfCompliance.PdfA2u,
                "pdfa4" => PdfCompliance.PdfA4,
                "pdfua1" => PdfCompliance.PdfUa1,
                _ => throw new ArgumentException(
                    $"Unknown PDF compliance: {p.PdfCompliance}. Supported: Pdf17, PdfA1a, PdfA1b, PdfA2a, PdfA2u, PdfA4, PdfUa1")
            };

        if (!string.IsNullOrEmpty(p.PdfPassword))
            options.EncryptionDetails = new PdfEncryptionDetails(p.PdfPassword, p.PdfPassword);

        if (p.EmbedFonts)
            options.EmbedFullFonts = true;

        if (p.DownsampleDpi > 0)
        {
            options.DownsampleOptions.Resolution = p.DownsampleDpi;
            options.DownsampleOptions.DownsampleImages = true;
        }

        return options;
    }

    private static ConvertParameters ExtractConvertParameters(OperationParameters parameters)
    {
        return new ConvertParameters(
            parameters.GetOptional<string?>("path"),
            parameters.GetOptional<string?>("sessionId"),
            parameters.GetOptional<string?>("outputPath"),
            parameters.GetOptional<string?>("format"),
            parameters.GetOptional<string?>("pdfCompliance"),
            parameters.GetOptional<string?>("pdfPassword"),
            parameters.GetOptional("embedFonts", false),
            parameters.GetOptional("downsampleDpi", 0));
    }

    /// <summary>
    ///     Parameters for the convert operation.
    /// </summary>
    /// <param name="Path">The input file path.</param>
    /// <param name="SessionId">The session ID for document from session.</param>
    /// <param name="OutputPath">The output file path.</param>
    /// <param name="Format">The output format.</param>
    /// <param name="PdfCompliance">The PDF compliance standard.</param>
    /// <param name="PdfPassword">The password for the output PDF.</param>
    /// <param name="EmbedFonts">Whether to embed all fonts in the PDF.</param>
    /// <param name="DownsampleDpi">The DPI for downsampling images in PDF output.</param>
    private sealed record ConvertParameters(
        string? Path,
        string? SessionId,
        string? OutputPath,
        string? Format,
        string? PdfCompliance,
        string? PdfPassword,
        bool EmbedFonts,
        int DownsampleDpi);
}
