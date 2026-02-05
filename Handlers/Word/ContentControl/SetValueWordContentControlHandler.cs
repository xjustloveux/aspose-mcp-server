using System.Globalization;
using Aspose.Words;
using Aspose.Words.Markup;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.ContentControl;

/// <summary>
///     Handler for setting the value of a content control in a Word document.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class SetValueWordContentControlHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "set_value";

    /// <summary>
    ///     Sets the value of a content control identified by index or tag.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: (index or tag) and value
    ///     - For PlainText/RichText: value is the text content
    ///     - For CheckBox: value is "true" or "false"
    ///     - For DropDownList/ComboBox: value is the item text to select
    ///     - For Date: value is a date string (ISO 8601 format)
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the value is invalid.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var index = parameters.GetOptional<int?>("index");
        var tag = parameters.GetOptional<string?>("tag");
        var value = parameters.GetRequired<string>("value");

        var doc = context.Document;
        var sdt = EditWordContentControlHandler.FindContentControl(doc, index, tag);

        if (sdt.LockContents)
            throw new InvalidOperationException("Content control contents are locked and cannot be modified");

        SetValue(sdt, value);

        MarkModified(context);

        var identifier = !string.IsNullOrEmpty(sdt.Tag) ? $"tag='{sdt.Tag}'" : $"index={index}";
        return new SuccessResult
        {
            Message = $"Content control ({sdt.SdtType}, {identifier}) value set to '{value}'."
        };
    }

    /// <summary>
    ///     Sets the value of a content control based on its type.
    /// </summary>
    /// <param name="sdt">The structured document tag.</param>
    /// <param name="value">The value to set.</param>
    /// <exception cref="ArgumentException">Thrown when the value is invalid for the content control type.</exception>
    private static void SetValue(StructuredDocumentTag sdt, string value)
    {
        switch (sdt.SdtType)
        {
            case SdtType.PlainText:
            case SdtType.RichText:
                SetTextValue(sdt, value);
                break;

            case SdtType.Checkbox:
                if (!bool.TryParse(value, out var isChecked))
                    throw new ArgumentException(
                        $"Invalid checkbox value: '{value}'. Use 'true' or 'false'.");
                sdt.Checked = isChecked;
                break;

            case SdtType.DropDownList:
            case SdtType.ComboBox:
                SelectListItem(sdt, value);
                break;

            case SdtType.Date:
                if (!DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dateValue))
                    throw new ArgumentException(
                        $"Invalid date value: '{value}'. Use ISO 8601 format (e.g., '2024-01-15').");
                sdt.FullDate = dateValue;
                break;

            default:
                SetTextValue(sdt, value);
                break;
        }
    }

    /// <summary>
    ///     Sets the text value of a content control by replacing its children.
    /// </summary>
    /// <param name="sdt">The structured document tag.</param>
    /// <param name="value">The text value to set.</param>
    private static void SetTextValue(StructuredDocumentTag sdt, string value)
    {
        sdt.RemoveAllChildren();
        sdt.AppendChild(new Run(sdt.Document, value));
    }

    /// <summary>
    ///     Selects a list item in a DropDownList or ComboBox content control.
    /// </summary>
    /// <param name="sdt">The structured document tag.</param>
    /// <param name="value">The item value or display text to select.</param>
    /// <exception cref="ArgumentException">Thrown when the item is not found in the list.</exception>
    private static void SelectListItem(StructuredDocumentTag sdt, string value)
    {
        var matchingItem = sdt.ListItems
            .FirstOrDefault(item =>
                string.Equals(item.Value, value, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(item.DisplayText, value, StringComparison.OrdinalIgnoreCase));

        if (matchingItem != null)
        {
            sdt.ListItems.SelectedValue = matchingItem;
            return;
        }

        var availableItems = string.Join(", ",
            sdt.ListItems.Select(i => i.Value));
        throw new ArgumentException(
            $"Item '{value}' not found in the list. Available items: {availableItems}");
    }
}
