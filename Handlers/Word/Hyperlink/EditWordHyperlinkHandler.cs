using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Hyperlink;

/// <summary>
///     Handler for editing hyperlinks in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractEditHyperlinkParameters(parameters);

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
        List<string> changes = [];

        if (!string.IsNullOrEmpty(p.Url))
        {
            WordHyperlinkHelper.ValidateUrlFormat(p.Url);
            hyperlinkField.Address = p.Url;
            changes.Add($"URL: {p.Url}");
        }

        if (!string.IsNullOrEmpty(p.SubAddress))
        {
            hyperlinkField.SubAddress = p.SubAddress;
            changes.Add($"SubAddress: {p.SubAddress}");
        }

        if (!string.IsNullOrEmpty(p.DisplayText))
        {
            hyperlinkField.Result = p.DisplayText;
            changes.Add($"Display text: {p.DisplayText}");
        }

        if (!string.IsNullOrEmpty(p.Tooltip))
        {
            hyperlinkField.ScreenTip = p.Tooltip;
            changes.Add($"Tooltip: {p.Tooltip}");
        }

        hyperlinkField.Update();

        MarkModified(context);

        var message = $"Hyperlink #{p.HyperlinkIndex} edited successfully\n";
        if (changes.Count > 0)
            message += $"Changes: {string.Join(", ", changes)}";
        else
            message += "No change parameters provided";

        return new SuccessResult { Message = message };
    }

    /// <summary>
    ///     Extracts edit hyperlink parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted edit hyperlink parameters.</returns>
    private static EditHyperlinkParameters ExtractEditHyperlinkParameters(OperationParameters parameters)
    {
        return new EditHyperlinkParameters(
            parameters.GetOptional("hyperlinkIndex", 0),
            parameters.GetOptional<string?>("url"),
            parameters.GetOptional<string?>("subAddress"),
            parameters.GetOptional<string?>("displayText"),
            parameters.GetOptional<string?>("tooltip")
        );
    }

    /// <summary>
    ///     Record to hold edit hyperlink parameters.
    /// </summary>
    private sealed record EditHyperlinkParameters(
        int HyperlinkIndex,
        string? Url,
        string? SubAddress,
        string? DisplayText,
        string? Tooltip);
}
