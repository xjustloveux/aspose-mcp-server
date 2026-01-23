using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Pdf.FileOperations;

/// <summary>
///     Handler for splitting a PDF document into multiple files.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var splitParams = ExtractSplitParameters(parameters);

        SecurityHelper.ValidateFilePath(splitParams.OutputDir, "outputDir", true);

        if (splitParams.PagesPerFile < 1 || splitParams.PagesPerFile > 1000)
            throw new ArgumentException("pagesPerFile must be between 1 and 1000");

        Directory.CreateDirectory(splitParams.OutputDir);

        var document = context.Document;
        var totalPages = document.Pages.Count;

        var baseName = splitParams.FileBaseName ?? (context.SourcePath != null
            ? SecurityHelper.SanitizeFileName(Path.GetFileNameWithoutExtension(context.SourcePath))
            : "document");

        var actualStartPage = splitParams.StartPage ?? 1;
        var actualEndPage = splitParams.EndPage ?? totalPages;

        if (actualStartPage < 1 || actualStartPage > totalPages)
            throw new ArgumentException($"startPage must be between 1 and {totalPages}");
        if (actualEndPage < actualStartPage || actualEndPage > totalPages)
            throw new ArgumentException($"endPage must be between {actualStartPage} and {totalPages}");

        if (splitParams.StartPage.HasValue || splitParams.EndPage.HasValue)
        {
            using var newDocument = new Document();
            for (var pageNum = actualStartPage; pageNum <= actualEndPage; pageNum++)
                newDocument.Pages.Add(document.Pages[pageNum]);

            var safeFileName =
                SecurityHelper.SanitizeFileName($"{baseName}_pages_{actualStartPage}-{actualEndPage}.pdf");
            var splitOutputPath = Path.Combine(splitParams.OutputDir, safeFileName);
            newDocument.Save(splitOutputPath);

            return new SuccessResult
            {
                Message =
                    $"PDF extracted pages {actualStartPage}-{actualEndPage} ({actualEndPage - actualStartPage + 1} pages). Output: {splitOutputPath}"
            };
        }

        var fileCount = 0;
        for (var i = 0; i < totalPages; i += splitParams.PagesPerFile)
        {
            using var newDocument = new Document();
            for (var j = 0; j < splitParams.PagesPerFile && i + j < totalPages; j++)
                newDocument.Pages.Add(document.Pages[i + j + 1]);

            var safeFileName = SecurityHelper.SanitizeFileName($"{baseName}_part_{++fileCount}.pdf");
            var splitOutputPath = Path.Combine(splitParams.OutputDir, safeFileName);
            newDocument.Save(splitOutputPath);
        }

        return new SuccessResult { Message = $"PDF split into {fileCount} files. Output: {splitParams.OutputDir}" };
    }

    /// <summary>
    ///     Extracts split parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted split parameters.</returns>
    private static SplitParameters ExtractSplitParameters(OperationParameters parameters)
    {
        return new SplitParameters(
            parameters.GetRequired<string>("outputDir"),
            parameters.GetOptional("pagesPerFile", 1),
            parameters.GetOptional<int?>("startPage"),
            parameters.GetOptional<int?>("endPage"),
            parameters.GetOptional<string?>("fileBaseName")
        );
    }

    /// <summary>
    ///     Record to hold split parameters.
    /// </summary>
    private sealed record SplitParameters(
        string OutputDir,
        int PagesPerFile,
        int? StartPage,
        int? EndPage,
        string? FileBaseName);
}
