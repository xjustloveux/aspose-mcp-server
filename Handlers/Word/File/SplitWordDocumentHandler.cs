using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;

namespace AsposeMcpServer.Handlers.Word.File;

/// <summary>
///     Handler for splitting Word documents.
/// </summary>
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
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var path = parameters.GetOptional<string?>("path");
        var sessionId = parameters.GetOptional<string?>("sessionId");
        var outputDir = parameters.GetOptional<string?>("outputDir");
        var splitBy = parameters.GetOptional("splitBy", "section");

        if (string.IsNullOrEmpty(path) && string.IsNullOrEmpty(sessionId))
            throw new ArgumentException("Either path or sessionId is required for split operation");
        if (string.IsNullOrEmpty(outputDir))
            throw new ArgumentException("outputDir is required for split operation");

        SecurityHelper.ValidateFilePath(outputDir, "outputDir", true);
        Directory.CreateDirectory(outputDir);

        Document doc;
        string fileBaseName;

        if (!string.IsNullOrEmpty(sessionId))
        {
            if (context.SessionManager == null)
                throw new InvalidOperationException("Session management is not enabled");

            var identity = context.IdentityAccessor?.GetCurrentIdentity() ?? SessionIdentity.GetAnonymous();
            doc = context.SessionManager.GetDocument<Document>(sessionId, identity);
            fileBaseName = $"session_{sessionId}";
        }
        else
        {
            SecurityHelper.ValidateFilePath(path!, allowAbsolutePaths: true);
            doc = new Document(path);
            fileBaseName = SecurityHelper.SanitizeFileName(Path.GetFileNameWithoutExtension(path!));
        }

        if (splitBy.ToLower() == "section")
        {
            for (var i = 0; i < doc.Sections.Count; i++)
            {
                var sectionDoc = new Document();
                sectionDoc.RemoveAllChildren();
                sectionDoc.AppendChild(sectionDoc.ImportNode(doc.Sections[i], true));

                var output = Path.Combine(outputDir, $"{fileBaseName}_section_{i + 1}.docx");
                sectionDoc.Save(output);
            }

            return $"Document split into {doc.Sections.Count} sections in: {outputDir}";
        }

        doc.UpdatePageLayout();

        var pageCount = doc.PageCount;
        for (var i = 0; i < pageCount; i++)
        {
            var pageDoc = doc.ExtractPages(i, 1);
            var output = Path.Combine(outputDir, $"{fileBaseName}_page_{i + 1}.docx");
            pageDoc.Save(output);
        }

        return $"Document split into {pageCount} pages in: {outputDir}";
    }
}
