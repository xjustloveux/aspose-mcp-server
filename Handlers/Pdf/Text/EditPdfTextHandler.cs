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
        var p = ExtractEditParameters(parameters);

        if (string.IsNullOrEmpty(p.OldText))
            throw new ArgumentException("oldText is required for edit operation");
        if (string.IsNullOrEmpty(p.NewText))
            throw new ArgumentException("newText is required for edit operation");

        var document = context.Document;
        if (p.PageIndex < 1 || p.PageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[p.PageIndex];

        var textFragmentAbsorber = new TextFragmentAbsorber(p.OldText);
        page.Accept(textFragmentAbsorber);

        var fragments = textFragmentAbsorber.TextFragments;
        var normalizedOldText =
            Regex.Replace(p.OldText, @"\s+", " ", RegexOptions.None, TimeSpan.FromSeconds(5)).Trim();

        if (fragments.Count == 0 && normalizedOldText != p.OldText)
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
                $"Text '{p.OldText}' not found on page {p.PageIndex}. Page text preview: {preview}");
        }

        var finalReplaceCount = p.ReplaceAll ? fragments.Count : 1;
        var replacedCount = 0;

        foreach (var fragment in fragments)
        {
            if (replacedCount >= finalReplaceCount)
                break;
            fragment.Text = p.NewText;
            replacedCount++;
        }

        MarkModified(context);

        return Success($"Replaced {replacedCount} occurrence(s) of '{p.OldText}' on page {p.PageIndex}.");
    }

    /// <summary>
    ///     Extracts parameters for edit operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static EditParameters ExtractEditParameters(OperationParameters parameters)
    {
        return new EditParameters(
            parameters.GetRequired<string>("oldText"),
            parameters.GetRequired<string>("newText"),
            parameters.GetOptional("pageIndex", 1),
            parameters.GetOptional("replaceAll", false)
        );
    }

    /// <summary>
    ///     Parameters for edit operation.
    /// </summary>
    /// <param name="OldText">The text to find.</param>
    /// <param name="NewText">The replacement text.</param>
    /// <param name="PageIndex">The 1-based page index.</param>
    /// <param name="ReplaceAll">Whether to replace all occurrences.</param>
    private sealed record EditParameters(string OldText, string NewText, int PageIndex, bool ReplaceAll);
}
