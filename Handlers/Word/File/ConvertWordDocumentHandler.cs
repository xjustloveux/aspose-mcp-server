using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;

namespace AsposeMcpServer.Handlers.Word.File;

/// <summary>
///     Handler for converting Word documents to other formats.
/// </summary>
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
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var path = parameters.GetOptional<string?>("path");
        var sessionId = parameters.GetOptional<string?>("sessionId");
        var outputPath = parameters.GetOptional<string?>("outputPath");
        var format = parameters.GetOptional<string?>("format");

        if (string.IsNullOrEmpty(path) && string.IsNullOrEmpty(sessionId))
            throw new ArgumentException("Either path or sessionId is required for convert operation");
        if (string.IsNullOrEmpty(outputPath))
            throw new ArgumentException("outputPath is required for convert operation");

        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        var outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir))
            Directory.CreateDirectory(outputDir);

        var formatLower = format?.ToLower();
        if (string.IsNullOrEmpty(formatLower))
        {
            var extension = Path.GetExtension(outputPath).TrimStart('.').ToLower();
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

        if (!string.IsNullOrEmpty(sessionId))
        {
            if (context.SessionManager == null)
                throw new InvalidOperationException("Session management is not enabled");

            var identity = context.IdentityAccessor?.GetCurrentIdentity() ?? SessionIdentity.GetAnonymous();
            doc = context.SessionManager.GetDocument<Document>(sessionId, identity);
            sourceDescription = $"session {sessionId}";
        }
        else
        {
            SecurityHelper.ValidateFilePath(path!, allowAbsolutePaths: true);
            doc = new Document(path);
            sourceDescription = path!;
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
            _ => throw new ArgumentException($"Unsupported format: {format}")
        };

        doc.Save(outputPath, saveFormat);
        return $"Document converted from {sourceDescription} to {outputPath} ({formatLower})";
    }
}
