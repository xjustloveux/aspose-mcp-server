using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;
using ModelContextProtocol;

namespace AsposeMcpServer.Handlers.Pdf.FileOperations;

/// <summary>
///     Handler for splitting a PDF document into multiple files.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class SplitPdfFileHandler : OperationHandlerBase<Document>
{
    /// <summary>
    ///     Maximum allowed total work units (totalPages × outputFileCount) per split call.
    ///     Applies only to full-document split mode (no explicit startPage/endPage), where the
    ///     caller controls <c>pagesPerFile</c> and therefore the number of output files.
    ///     A 100-page PDF split into 1 page per file = 100 × 100 = 10,000 units (well within
    ///     the cap); a 10,000-page PDF split into 1 page per file = 100M units (rejected).
    ///     The bound ensures that the total I/O load stays proportional to document size even
    ///     when pagesPerFile is small.
    /// </summary>
    private const int MaxTotalWorkUnits = 100_000;

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
    /// <exception cref="ArgumentException">
    ///     Thrown when <c>pagesPerFile</c> is outside [1,1000]; when <c>startPage</c> or
    ///     <c>endPage</c> is out of range; when, in full-split mode (no explicit page range),
    ///     the product of totalPages and output file count exceeds <see cref="MaxTotalWorkUnits" />;
    ///     or when path validation fails.
    /// </exception>
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
            // H27: resolve symlinks immediately before the sink (bug 20260415-symlink-toctou-sweep).
            splitOutputPath = SecurityHelper.ResolveAndEnsureWithinAllowlist(splitOutputPath,
                context.ServerConfig?.AllowedBasePaths ?? [], nameof(splitOutputPath));
            newDocument.Save(splitOutputPath);

            return new SuccessResult
            {
                Message =
                    $"PDF extracted pages {actualStartPage}-{actualEndPage} ({actualEndPage - actualStartPage + 1} pages). Output: {splitOutputPath}"
            };
        }

        var fileCount = 0;
        var totalSplits = (int)Math.Ceiling((double)totalPages / splitParams.PagesPerFile);

        // Orthogonal DoS guard: even with pagesPerFile ≥ 1, a very large document combined
        // with a small pagesPerFile creates excessive I/O (e.g., 10,000 pages × 10,000 files).
        var totalWorkUnits = (long)totalPages * totalSplits;
        if (totalWorkUnits > MaxTotalWorkUnits)
            throw new ArgumentException(
                $"Split would require {totalWorkUnits} work units (totalPages × outputFiles = {totalPages} × {totalSplits}). " +
                $"Maximum allowed is {MaxTotalWorkUnits}. Increase pagesPerFile to reduce the number of output files.");

        for (var i = 0; i < totalPages; i += splitParams.PagesPerFile)
        {
            using var newDocument = new Document();
            for (var j = 0; j < splitParams.PagesPerFile && i + j < totalPages; j++)
                newDocument.Pages.Add(document.Pages[i + j + 1]);

            var safeFileName = SecurityHelper.SanitizeFileName($"{baseName}_part_{++fileCount}.pdf");
            var splitOutputPath = Path.Combine(splitParams.OutputDir, safeFileName);
            // H27: resolve symlinks immediately before the sink (bug 20260415-symlink-toctou-sweep).
            splitOutputPath = SecurityHelper.ResolveAndEnsureWithinAllowlist(splitOutputPath,
                context.ServerConfig?.AllowedBasePaths ?? [], nameof(splitOutputPath));
            newDocument.Save(splitOutputPath);

            var splitProgress = fileCount * 100 / totalSplits;
            context.Progress?.Report(new ProgressNotificationValue
            {
                Progress = splitProgress,
                Total = 100,
                Message = $"Created split file {fileCount} of {totalSplits}"
            });
        }

        context.Progress?.Report(new ProgressNotificationValue
            { Progress = 100, Total = 100, Message = "Split completed" });

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
