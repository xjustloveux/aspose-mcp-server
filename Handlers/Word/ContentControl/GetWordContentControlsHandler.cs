using Aspose.Words;
using Aspose.Words.Markup;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Word.ContentControl;

namespace AsposeMcpServer.Handlers.Word.ContentControl;

/// <summary>
///     Handler for getting content controls from a Word document.
/// </summary>
[ResultType(typeof(GetContentControlsResult))]
public class GetWordContentControlsHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets all content controls from the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: tag (filter by tag), type (filter by type)
    /// </param>
    /// <returns>A result containing the list of content controls.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var filterTag = parameters.GetOptional<string?>("tag");
        var filterType = parameters.GetOptional<string?>("type");

        var doc = context.Document;
        var sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);

        var contentControls = new List<ContentControlInfo>();
        var index = 0;

        foreach (var sdt in sdtNodes.Cast<StructuredDocumentTag>())
        {
            if (!string.IsNullOrEmpty(filterTag) &&
                !string.Equals(sdt.Tag, filterTag, StringComparison.OrdinalIgnoreCase))
            {
                index++;
                continue;
            }

            if (!string.IsNullOrEmpty(filterType) &&
                !string.Equals(sdt.SdtType.ToString(), filterType, StringComparison.OrdinalIgnoreCase))
            {
                index++;
                continue;
            }

            contentControls.Add(new ContentControlInfo
            {
                Index = index,
                Tag = sdt.Tag,
                Title = sdt.Title,
                Type = sdt.SdtType.ToString(),
                Value = GetContentControlValue(sdt),
                Placeholder = sdt.IsShowingPlaceholderText ? sdt.GetText().Trim() : null,
                LockContents = sdt.LockContents,
                LockDeletion = sdt.LockContentControl
            });

            index++;
        }

        return new GetContentControlsResult
        {
            Count = contentControls.Count,
            ContentControls = contentControls,
            Message = contentControls.Count == 0 ? "No content controls found in the document." : null
        };
    }

    /// <summary>
    ///     Gets the value of a content control based on its type.
    /// </summary>
    /// <param name="sdt">The structured document tag.</param>
    /// <returns>The string value of the content control.</returns>
    private static string? GetContentControlValue(StructuredDocumentTag sdt)
    {
        if (sdt.IsShowingPlaceholderText)
            return null;

        return sdt.SdtType switch
        {
            SdtType.Checkbox => sdt.Checked.ToString(),
            SdtType.Date => sdt.FullDate.ToString("o"),
            SdtType.DropDownList or SdtType.ComboBox => sdt.ListItems.SelectedValue?.Value,
            _ => sdt.GetText().Trim('\r', '\n', '\a')
        };
    }
}
