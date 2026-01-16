using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Word.Styles;

/// <summary>
///     Handler for copying styles from one Word document to another.
/// </summary>
public class CopyWordStylesHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "copy_styles";

    /// <summary>
    ///     Copies styles from source document to target document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: sourceDocument
    ///     Optional: styleNames, overwriteExisting
    /// </param>
    /// <returns>Success message with copy details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractCopyWordStylesParameters(parameters);

        if (string.IsNullOrEmpty(p.SourceDocument))
            throw new ArgumentException("sourceDocument is required for copy_styles operation");

        SecurityHelper.ValidateFilePath(p.SourceDocument, "sourceDocument", true);

        if (!System.IO.File.Exists(p.SourceDocument))
            throw new FileNotFoundException($"Source document not found: {p.SourceDocument}");

        var targetDoc = context.Document;
        var sourceDoc = new Document(p.SourceDocument);

        var styleNamesList = p.StyleNames?.ToList() ?? [];
        var copyAll = styleNamesList.Count == 0;
        var copiedCount = 0;
        var skippedCount = 0;

        foreach (var sourceStyle in sourceDoc.Styles)
        {
            if (!copyAll && !styleNamesList.Contains(sourceStyle.Name))
                continue;

            var existingStyle = targetDoc.Styles[sourceStyle.Name];

            if (existingStyle != null && !p.OverwriteExisting)
            {
                skippedCount++;
                continue;
            }

            try
            {
                if (existingStyle != null && p.OverwriteExisting)
                {
                    WordStyleHelper.CopyStyleProperties(sourceStyle, existingStyle);
                }
                else
                {
                    var newStyle = targetDoc.Styles.Add(sourceStyle.Type, sourceStyle.Name);
                    WordStyleHelper.CopyStyleProperties(sourceStyle, newStyle);
                }

                copiedCount++;
            }
            catch (Exception ex)
            {
                skippedCount++;
                Console.Error.WriteLine($"[WARN] Failed to copy style '{sourceStyle.Name}': {ex.Message}");
            }
        }

        MarkModified(context);

        return Success(
            $"Copied {copiedCount} style(s) from {Path.GetFileName(p.SourceDocument)}. Skipped: {skippedCount}.");
    }

    private static CopyWordStylesParameters ExtractCopyWordStylesParameters(OperationParameters parameters)
    {
        return new CopyWordStylesParameters(
            parameters.GetRequired<string>("sourceDocument"),
            parameters.GetOptional<string[]?>("styleNames"),
            parameters.GetOptional("overwriteExisting", false));
    }

    private record CopyWordStylesParameters(
        string SourceDocument,
        string[]? StyleNames,
        bool OverwriteExisting);
}
