using System.Text.RegularExpressions;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Text;

/// <summary>
///     Handler for editing (replacing) text in PDF documents.
/// </summary>
public class EditPdfTextHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "edit";

    /// <summary>
    ///     Edits (replaces) text on the specified page of the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: oldText, newText
    ///     Optional: pageIndex, replaceAll
    /// </param>
    /// <returns>Success message with replacement count.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var oldText = parameters.GetRequired<string>("oldText");
        var newText = parameters.GetRequired<string>("newText");
        var pageIndex = parameters.GetOptional("pageIndex", 1);
        var replaceAll = parameters.GetOptional("replaceAll", false);

        if (string.IsNullOrEmpty(oldText))
            throw new ArgumentException("oldText is required for edit operation");
        if (string.IsNullOrEmpty(newText))
            throw new ArgumentException("newText is required for edit operation");

        var document = context.Document;
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];

        var textFragmentAbsorber = new TextFragmentAbsorber(oldText);
        page.Accept(textFragmentAbsorber);

        var fragments = textFragmentAbsorber.TextFragments;
        var normalizedOldText = Regex.Replace(oldText, @"\s+", " ").Trim();

        if (fragments.Count == 0 && normalizedOldText != oldText)
        {
            textFragmentAbsorber = new TextFragmentAbsorber(normalizedOldText);
            page.Accept(textFragmentAbsorber);
            fragments = textFragmentAbsorber.TextFragments;
        }

        if (fragments.Count == 0)
        {
            var textAbsorber = new TextAbsorber();
            page.Accept(textAbsorber);
            var pageText = textAbsorber.Text ?? "";
            var preview = pageText.Length > 200 ? pageText[..200] + "..." : pageText;
            throw new ArgumentException(
                $"Text '{oldText}' not found on page {pageIndex}. Page text preview: {preview}");
        }

        var finalReplaceCount = replaceAll ? fragments.Count : 1;
        var replacedCount = 0;

        foreach (var fragment in fragments)
        {
            if (replacedCount >= finalReplaceCount)
                break;
            fragment.Text = newText;
            replacedCount++;
        }

        MarkModified(context);

        return Success($"Replaced {replacedCount} occurrence(s) of '{oldText}' on page {pageIndex}.");
    }
}
