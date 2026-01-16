using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Hyperlink;

/// <summary>
///     Handler for deleting hyperlinks from Word documents.
/// </summary>
public class DeleteWordHyperlinkHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes a hyperlink from the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: hyperlinkIndex
    ///     Optional: keepText (default: false)
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractDeleteHyperlinkParameters(parameters);

        var doc = context.Document;
        var hyperlinkFields = WordHyperlinkHelper.GetAllHyperlinks(doc);

        if (p.HyperlinkIndex < 0 || p.HyperlinkIndex >= hyperlinkFields.Count)
        {
            var availableInfo = hyperlinkFields.Count > 0
                ? $" (valid index: 0-{hyperlinkFields.Count - 1})"
                : " (document has no hyperlinks)";
            throw new ArgumentException(
                $"Hyperlink index {p.HyperlinkIndex} is out of range (document has {hyperlinkFields.Count} hyperlinks){availableInfo}. Use get operation to view all available hyperlinks");
        }

        var hyperlinkField = hyperlinkFields[p.HyperlinkIndex];
        var displayTextValue = hyperlinkField.Result ?? "";
        var address = hyperlinkField.Address ?? "";

        if (p.KeepText)
            hyperlinkField.Unlink();
        else
            hyperlinkField.Remove();

        MarkModified(context);

        var remainingCount = WordHyperlinkHelper.GetAllHyperlinks(doc).Count;

        var result = $"Hyperlink #{p.HyperlinkIndex} deleted successfully\n";
        result += $"Display text: {displayTextValue}\n";
        result += $"Address: {address}\n";
        result += $"Keep text: {(p.KeepText ? "Yes (unlinked)" : "No (removed)")}\n";
        result += $"Remaining hyperlinks in document: {remainingCount}";

        return result;
    }

    /// <summary>
    ///     Extracts delete hyperlink parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted delete hyperlink parameters.</returns>
    private static DeleteHyperlinkParameters ExtractDeleteHyperlinkParameters(OperationParameters parameters)
    {
        return new DeleteHyperlinkParameters(
            parameters.GetOptional("hyperlinkIndex", 0),
            parameters.GetOptional("keepText", false)
        );
    }

    /// <summary>
    ///     Record to hold delete hyperlink parameters.
    /// </summary>
    private sealed record DeleteHyperlinkParameters(int HyperlinkIndex, bool KeepText);
}
