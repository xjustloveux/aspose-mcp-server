using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Field;

/// <summary>
///     Handler for adding form fields to Word documents.
/// </summary>
public class AddFormFieldWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add_form_field";

    /// <summary>
    ///     Adds a form field (text input, checkbox, or dropdown) to the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: formFieldType (TextInput, CheckBox, DropDown), fieldName
    ///     Optional: defaultValue, options, checkedValue
    /// </param>
    /// <returns>Success message with form field details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var formFieldType = parameters.GetOptional<string?>("formFieldType");
        var fieldName = parameters.GetOptional<string?>("fieldName");
        var defaultValue = parameters.GetOptional<string?>("defaultValue");
        var options = parameters.GetOptional<string[]?>("options");
        var checkedValue = parameters.GetOptional<bool?>("checkedValue");

        if (string.IsNullOrEmpty(formFieldType))
            throw new ArgumentException("formFieldType is required for add_form_field operation");
        if (string.IsNullOrEmpty(fieldName))
            throw new ArgumentException("fieldName is required for add_form_field operation");

        var document = context.Document;
        var builder = new DocumentBuilder(document);
        builder.MoveToDocumentEnd();

        switch (formFieldType.ToLower())
        {
            case "textinput":
                builder.InsertTextInput(fieldName, TextFormFieldType.Regular, "", defaultValue ?? "", 0);
                break;
            case "checkbox":
                builder.InsertCheckBox(fieldName, checkedValue ?? false, 0);
                break;
            case "dropdown":
                if (options == null || options.Length == 0)
                    throw new ArgumentException("options array is required for DropDown type");
                builder.InsertComboBox(fieldName, options, 0);
                break;
            default:
                throw new ArgumentException(
                    $"Invalid formFieldType: {formFieldType}. Must be 'TextInput', 'CheckBox', or 'DropDown'");
        }

        MarkModified(context);
        return Success($"{formFieldType} field '{fieldName}' added");
    }
}
