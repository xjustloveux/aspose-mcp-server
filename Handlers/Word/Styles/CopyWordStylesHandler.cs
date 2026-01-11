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
        var sourceDocument = parameters.GetRequired<string>("sourceDocument");
        var styleNames = parameters.GetOptional<string[]?>("styleNames");
        var overwriteExisting = parameters.GetOptional("overwriteExisting", false);

        if (string.IsNullOrEmpty(sourceDocument))
            throw new ArgumentException("sourceDocument is required for copy_styles operation");

        SecurityHelper.ValidateFilePath(sourceDocument, "sourceDocument", true);

        if (!System.IO.File.Exists(sourceDocument))
            throw new FileNotFoundException($"Source document not found: {sourceDocument}");

        var targetDoc = context.Document;
        var sourceDoc = new Document(sourceDocument);

        var styleNamesList = styleNames?.ToList() ?? [];
        var copyAll = styleNamesList.Count == 0;
        var copiedCount = 0;
        var skippedCount = 0;

        foreach (var sourceStyle in sourceDoc.Styles)
        {
            if (!copyAll && !styleNamesList.Contains(sourceStyle.Name))
                continue;

            var existingStyle = targetDoc.Styles[sourceStyle.Name];

            if (existingStyle != null && !overwriteExisting)
            {
                skippedCount++;
                continue;
            }

            try
            {
                if (existingStyle != null && overwriteExisting)
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
            $"Copied {copiedCount} style(s) from {Path.GetFileName(sourceDocument)}. Skipped: {skippedCount}.");
    }
}
