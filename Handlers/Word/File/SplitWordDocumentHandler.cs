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
    /// <exception cref="ArgumentException">Thrown when required parameters are missing.</exception>
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
                sectionDoc.Save(output);
            }

            return new SuccessResult
                { Message = $"Document split into {doc.Sections.Count} sections in: {p.OutputDir}" };
        }

        doc.UpdatePageLayout();

        var pageCount = doc.PageCount;
        for (var i = 0; i < pageCount; i++)
        {
            var pageDoc = doc.ExtractPages(i, 1);
            var output = Path.Combine(p.OutputDir, $"{fileBaseName}_page_{i + 1}.docx");
            pageDoc.Save(output);
        }

        return new SuccessResult { Message = $"Document split into {pageCount} pages in: {p.OutputDir}" };
    }

    private static SplitParameters ExtractSplitParameters(OperationParameters parameters)
    {
        return new SplitParameters(
            parameters.GetOptional<string?>("path"),
            parameters.GetOptional<string?>("sessionId"),
            parameters.GetOptional<string?>("outputDir"),
            parameters.GetOptional("splitBy", "section"));
    }

    private sealed record SplitParameters(
        string? Path,
        string? SessionId,
        string? OutputDir,
        string SplitBy);
}
