using Aspose.Words;
using Aspose.Words.Saving;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Progress;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

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

        if (formatLower == "pdf" && context.Progress != null)
        {
            var pdfSaveOptions = new PdfSaveOptions
            {
                ProgressCallback = new WordsProgressAdapter(context.Progress)
            };
            doc.Save(p.OutputPath, pdfSaveOptions);
        }
        else
        {
            doc.Save(p.OutputPath, saveFormat);
        }

        return new SuccessResult
            { Message = $"Document converted from {sourceDescription} to {p.OutputPath} ({formatLower})" };
    }

    private static ConvertParameters ExtractConvertParameters(OperationParameters parameters)
    {
        return new ConvertParameters(
            parameters.GetOptional<string?>("path"),
            parameters.GetOptional<string?>("sessionId"),
            parameters.GetOptional<string?>("outputPath"),
            parameters.GetOptional<string?>("format"));
    }

    private sealed record ConvertParameters(
        string? Path,
        string? SessionId,
        string? OutputPath,
        string? Format);
}
