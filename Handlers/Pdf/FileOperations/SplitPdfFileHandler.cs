using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Pdf.FileOperations;

/// <summary>
///     Handler for splitting a PDF document into multiple files.
/// </summary>
public class SplitPdfFileHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "split";

    /// <summary>
    ///     Splits a PDF document into multiple files.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: outputDir
    ///     Optional: pagesPerFile (default: 1), startPage, endPage, fileBaseName
    /// </param>
    /// <returns>Success message with split result.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var outputDir = parameters.GetRequired<string>("outputDir");
        var pagesPerFile = parameters.GetOptional("pagesPerFile", 1);
        var startPage = parameters.GetOptional<int?>("startPage");
        var endPage = parameters.GetOptional<int?>("endPage");
        var fileBaseName = parameters.GetOptional<string?>("fileBaseName");

        SecurityHelper.ValidateFilePath(outputDir, "outputDir", true);

        if (pagesPerFile < 1 || pagesPerFile > 1000)
            throw new ArgumentException("pagesPerFile must be between 1 and 1000");

        Directory.CreateDirectory(outputDir);

        var document = context.Document;
        var totalPages = document.Pages.Count;

        var baseName = fileBaseName ?? (context.SourcePath != null
            ? SecurityHelper.SanitizeFileName(Path.GetFileNameWithoutExtension(context.SourcePath))
            : "document");

        var actualStartPage = startPage ?? 1;
        var actualEndPage = endPage ?? totalPages;

        if (actualStartPage < 1 || actualStartPage > totalPages)
            throw new ArgumentException($"startPage must be between 1 and {totalPages}");
        if (actualEndPage < actualStartPage || actualEndPage > totalPages)
            throw new ArgumentException($"endPage must be between {actualStartPage} and {totalPages}");

        if (startPage.HasValue || endPage.HasValue)
        {
            using var newDocument = new Document();
            for (var pageNum = actualStartPage; pageNum <= actualEndPage; pageNum++)
                newDocument.Pages.Add(document.Pages[pageNum]);

            var safeFileName =
                SecurityHelper.SanitizeFileName($"{baseName}_pages_{actualStartPage}-{actualEndPage}.pdf");
            var splitOutputPath = Path.Combine(outputDir, safeFileName);
            newDocument.Save(splitOutputPath);

            return Success(
                $"PDF extracted pages {actualStartPage}-{actualEndPage} ({actualEndPage - actualStartPage + 1} pages). Output: {splitOutputPath}");
        }

        var fileCount = 0;
        for (var i = 0; i < totalPages; i += pagesPerFile)
        {
            using var newDocument = new Document();
            for (var j = 0; j < pagesPerFile && i + j < totalPages; j++)
                newDocument.Pages.Add(document.Pages[i + j + 1]);

            var safeFileName = SecurityHelper.SanitizeFileName($"{baseName}_part_{++fileCount}.pdf");
            var splitOutputPath = Path.Combine(outputDir, safeFileName);
            newDocument.Save(splitOutputPath);
        }

        return Success($"PDF split into {fileCount} files. Output: {outputDir}");
    }
}
