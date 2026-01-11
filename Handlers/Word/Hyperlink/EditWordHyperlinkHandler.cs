using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Hyperlink;

/// <summary>
///     Handler for editing hyperlinks in Word documents.
/// </summary>
public class EditWordHyperlinkHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "edit";

    /// <summary>
    ///     Edits an existing hyperlink.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: hyperlinkIndex
    ///     Optional: url, subAddress, displayText, tooltip
    /// </param>
    /// <returns>Success message with edit details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var hyperlinkIndex = parameters.GetOptional("hyperlinkIndex", 0);
        var url = parameters.GetOptional<string?>("url");
        var subAddress = parameters.GetOptional<string?>("subAddress");
        var displayText = parameters.GetOptional<string?>("displayText");
        var tooltip = parameters.GetOptional<string?>("tooltip");

        var doc = context.Document;
        var hyperlinkFields = WordHyperlinkHelper.GetAllHyperlinks(doc);

        if (hyperlinkIndex < 0 || hyperlinkIndex >= hyperlinkFields.Count)
        {
            var availableInfo = hyperlinkFields.Count > 0
                ? $" (valid index: 0-{hyperlinkFields.Count - 1})"
                : " (document has no hyperlinks)";
            throw new ArgumentException(
                $"Hyperlink index {hyperlinkIndex} is out of range (document has {hyperlinkFields.Count} hyperlinks){availableInfo}. Use get operation to view all available hyperlinks");
        }

        var hyperlinkField = hyperlinkFields[hyperlinkIndex];
        List<string> changes = [];

        if (!string.IsNullOrEmpty(url))
        {
            WordHyperlinkHelper.ValidateUrlFormat(url);
            hyperlinkField.Address = url;
            changes.Add($"URL: {url}");
        }

        if (!string.IsNullOrEmpty(subAddress))
        {
            hyperlinkField.SubAddress = subAddress;
            changes.Add($"SubAddress: {subAddress}");
        }

        if (!string.IsNullOrEmpty(displayText))
        {
            hyperlinkField.Result = displayText;
            changes.Add($"Display text: {displayText}");
        }

        if (!string.IsNullOrEmpty(tooltip))
        {
            hyperlinkField.ScreenTip = tooltip;
            changes.Add($"Tooltip: {tooltip}");
        }

        hyperlinkField.Update();

        MarkModified(context);

        var result = $"Hyperlink #{hyperlinkIndex} edited successfully\n";
        if (changes.Count > 0)
            result += $"Changes: {string.Join(", ", changes)}";
        else
            result += "No change parameters provided";

        return result;
    }
}
