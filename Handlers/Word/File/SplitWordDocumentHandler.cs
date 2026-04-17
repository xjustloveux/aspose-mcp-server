using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.File;

/// <summary>
///     Handler for splitting Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class SplitWordDocumentHandler : OperationHandlerBase<Document>
{
    /// <summary>
    ///     Maximum allowed total work units for page-split mode (totalPages × outputFileCount).
    ///     In page-split mode each page becomes exactly one output file, so work units =
    ///     pageCount × pageCount = pageCount².  A 316-page document hits ~100,000 units at
    ///     the limit.  Section-split mode is exempt because section count is structurally
    ///     bounded and each section produces exactly one file (work units = sectionCount × 1).
    /// </summary>
    private const int MaxTotalWorkUnits = 100_000;

    /// <inheritdoc />
    public override string Operation => "split";

    /// <summary>
    ///     Splits a Word document by sections or pages.
    /// </summary>
    /// <param name="context">The operation context.</param>
    /// <param name="parameters">
    ///     Required: outputDir, either path or sessionId
    ///     Optional: splitBy (default: section)
    /// </param>
    /// <returns>Success message with split details.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when required parameters are missing; or, in page-split mode, when the
    ///     square of the page count exceeds <see cref="MaxTotalWorkUnits" /> (DoS guard).
    /// </exception>
    /// <exception cref="InvalidOperationException">Thrown when session management is not enabled.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractSplitParameters(parameters);

        if (string.IsNullOrEmpty(p.Path) && string.IsNullOrEmpty(p.SessionId))
            throw new ArgumentException("Either path or sessionId is required for split operation");
        if (string.IsNullOrEmpty(p.OutputDir))
            throw new ArgumentException("outputDir is required for split operation");

        SecurityHelper.ValidateFilePath(p.OutputDir, "outputDir", true);
        Directory.CreateDirectory(p.OutputDir);

        Document doc;
        string fileBaseName;

        if (!string.IsNullOrEmpty(p.SessionId))
        {
            if (context.SessionManager == null)
                throw new InvalidOperationException("Session management is not enabled");

            var identity = context.IdentityAccessor?.GetCurrentIdentity() ?? SessionIdentity.GetAnonymous();
            doc = context.SessionManager.GetDocument<Document>(p.SessionId, identity);
            fileBaseName = $"session_{p.SessionId}";
        }
        else
        {
            SecurityHelper.ValidateFilePath(p.Path!, allowAbsolutePaths: true);
            doc = new Document(p.Path);
            fileBaseName = SecurityHelper.SanitizeFileName(Path.GetFileNameWithoutExtension(p.Path!));
        }

        if (string.Equals(p.SplitBy, "section", StringComparison.OrdinalIgnoreCase))
        {
            for (var i = 0; i < doc.Sections.Count; i++)
            {
                var sectionDoc = new Document();
                sectionDoc.RemoveAllChildren();
                sectionDoc.AppendChild(sectionDoc.ImportNode(doc.Sections[i], true));

                var output = Path.Combine(p.OutputDir, $"{fileBaseName}_section_{i + 1}.docx");
                // H5: resolve symlinks immediately before the sink (bug 20260415-symlink-toctou-sweep).
                output = SecurityHelper.ResolveAndEnsureWithinAllowlist(output,
                    context.ServerConfig?.AllowedBasePaths ?? [], nameof(output));
                sectionDoc.Save(output);
            }

            return new SuccessResult
                { Message = $"Document split into {doc.Sections.Count} sections in: {p.OutputDir}" };
        }

        doc.UpdatePageLayout();

        var pageCount = doc.PageCount;

        // Orthogonal DoS guard: in page-split mode each page produces one output file,
        // so the total work count equals the page count squared. A 316-page document
        // reaches the cap — larger documents must use section-split or a range-aware
        // tool instead.
        var totalWorkUnits = (long)pageCount * pageCount;
        if (totalWorkUnits > MaxTotalWorkUnits)
            throw new ArgumentException(
                $"Page-split would require {totalWorkUnits} work units (totalPages² = {pageCount}²). " +
                $"Maximum allowed is {MaxTotalWorkUnits}. Use section-split mode or a smaller document.");

        for (var i = 0; i < pageCount; i++)
        {
            var pageDoc = doc.ExtractPages(i, 1);
            var output = Path.Combine(p.OutputDir, $"{fileBaseName}_page_{i + 1}.docx");
            // H5: resolve symlinks immediately before each per-page sink (bug 20260415-symlink-toctou-sweep).
            output = SecurityHelper.ResolveAndEnsureWithinAllowlist(output,
                context.ServerConfig?.AllowedBasePaths ?? [], nameof(output));
            pageDoc.Save(output);
        }

        return new SuccessResult { Message = $"Document split into {pageCount} pages in: {p.OutputDir}" };
    }

    /// <summary>
    ///     Extracts split parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters bag; must not be null.</param>
    /// <returns>A <see cref="SplitParameters" /> record populated from the caller's input.</returns>
    private static SplitParameters ExtractSplitParameters(OperationParameters parameters)
    {
        return new SplitParameters(
            parameters.GetOptional<string?>("path"),
            parameters.GetOptional<string?>("sessionId"),
            parameters.GetOptional<string?>("outputDir"),
            parameters.GetOptional("splitBy", "section"));
    }

    /// <summary>
    ///     Holds the validated split parameters extracted from the caller's operation parameters.
    /// </summary>
    /// <param name="Path">Source document path; null when a session is used instead.</param>
    /// <param name="SessionId">Session identifier for session-backed documents; null when a path is supplied.</param>
    /// <param name="OutputDir">Directory into which output files are written; must not be null or empty.</param>
    /// <param name="SplitBy">
    ///     Split mode: <c>"section"</c> (default) splits by document section; any other value
    ///     triggers page-by-page split which requires <see cref="MaxTotalWorkUnits" /> guard.
    /// </param>
    private sealed record SplitParameters(
        string? Path,
        string? SessionId,
        string? OutputDir,
        string SplitBy);
}
