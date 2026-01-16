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
        var p = ExtractEditFormFieldParameters(parameters);

        var document = context.Document;
        var field = document.Range.FormFields[p.FieldName];

        if (field == null)
            throw new ArgumentException($"Form field '{p.FieldName}' not found");

        if (field.Type == FieldType.FieldFormTextInput && p.Value != null)
            field.Result = p.Value;
        else if (field.Type == FieldType.FieldFormCheckBox && p.CheckedValue.HasValue)
            field.Checked = p.CheckedValue.Value;
        else if (field.Type == FieldType.FieldFormDropDown &&
                 p is { SelectedIndex: >= 0 and var selectedIndex } &&
                 selectedIndex < field.DropDownItems.Count)
            field.DropDownSelectedIndex = selectedIndex;

        MarkModified(context);
        return Success($"Form field '{p.FieldName}' updated");
    }

    /// <summary>
    ///     Extracts and validates parameters for the edit form field operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    /// <exception cref="ArgumentException">Thrown when fieldName is not provided.</exception>
    private static EditFormFieldParameters ExtractEditFormFieldParameters(OperationParameters parameters)
    {
        var fieldName = parameters.GetOptional<string?>("fieldName");
        var value = parameters.GetOptional<string?>("value");
        var checkedValue = parameters.GetOptional<bool?>("checkedValue");
        var selectedIndex = parameters.GetOptional<int?>("selectedIndex");

        if (string.IsNullOrEmpty(fieldName))
            throw new ArgumentException("fieldName is required for edit_form_field operation");

        return new EditFormFieldParameters(fieldName, value, checkedValue, selectedIndex);
    }

    /// <summary>
    ///     Parameters for the edit form field operation.
    /// </summary>
    /// <param name="FieldName">The name of the form field to edit.</param>
    /// <param name="Value">The new value for text input fields.</param>
    /// <param name="CheckedValue">The new checked state for checkbox fields.</param>
    /// <param name="SelectedIndex">The new selected index for dropdown fields.</param>
    private sealed record EditFormFieldParameters(
        string FieldName,
        string? Value,
        bool? CheckedValue,
        int? SelectedIndex);
}
