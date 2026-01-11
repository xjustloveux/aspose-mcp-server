using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Field;

/// <summary>
///     Handler for editing form fields in Word documents.
/// </summary>
public class EditFormFieldWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "edit_form_field";

    /// <summary>
    ///     Edits an existing form field's value or state.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: fieldName
    ///     Optional: value (for TextInput), checkedValue (for CheckBox), selectedIndex (for DropDown)
    /// </param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var fieldName = parameters.GetOptional<string?>("fieldName");
        var value = parameters.GetOptional<string?>("value");
        var checkedValue = parameters.GetOptional<bool?>("checkedValue");
        var selectedIndex = parameters.GetOptional<int?>("selectedIndex");

        if (string.IsNullOrEmpty(fieldName))
            throw new ArgumentException("fieldName is required for edit_form_field operation");

        var document = context.Document;
        var field = document.Range.FormFields[fieldName];

        if (field == null)
            throw new ArgumentException($"Form field '{fieldName}' not found");

        if (field.Type == FieldType.FieldFormTextInput && value != null)
            field.Result = value;
        else if (field.Type == FieldType.FieldFormCheckBox && checkedValue.HasValue)
            field.Checked = checkedValue.Value;
        else if (field.Type == FieldType.FieldFormDropDown && selectedIndex.HasValue)
            if (selectedIndex.Value >= 0 && selectedIndex.Value < field.DropDownItems.Count)
                field.DropDownSelectedIndex = selectedIndex.Value;

        MarkModified(context);
        return Success($"Form field '{fieldName}' updated");
    }
}
