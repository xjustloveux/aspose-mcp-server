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
        var p = ExtractAddFormFieldParameters(parameters);

        var document = context.Document;
        var builder = new DocumentBuilder(document);
        builder.MoveToDocumentEnd();

        switch (p.FormFieldType.ToLower())
        {
            case "textinput":
                builder.InsertTextInput(p.FieldName, TextFormFieldType.Regular, "", p.DefaultValue ?? "", 0);
                break;
            case "checkbox":
                builder.InsertCheckBox(p.FieldName, p.CheckedValue ?? false, 0);
                break;
            case "dropdown":
                if (p.Options == null || p.Options.Length == 0)
                    throw new ArgumentException("options array is required for DropDown type");
                builder.InsertComboBox(p.FieldName, p.Options, 0);
                break;
            default:
                throw new ArgumentException(
                    $"Invalid formFieldType: {p.FormFieldType}. Must be 'TextInput', 'CheckBox', or 'DropDown'");
        }

        MarkModified(context);
        return Success($"{p.FormFieldType} field '{p.FieldName}' added");
    }

    /// <summary>
    ///     Extracts and validates parameters for the add form field operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are not provided.</exception>
    private static AddFormFieldParameters ExtractAddFormFieldParameters(OperationParameters parameters)
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

        return new AddFormFieldParameters(formFieldType, fieldName, defaultValue, options, checkedValue);
    }

    /// <summary>
    ///     Parameters for the add form field operation.
    /// </summary>
    /// <param name="FormFieldType">The type of form field (TextInput, CheckBox, DropDown).</param>
    /// <param name="FieldName">The name of the form field.</param>
    /// <param name="DefaultValue">The default value for text input fields.</param>
    /// <param name="Options">The options for dropdown fields.</param>
    /// <param name="CheckedValue">The initial checked state for checkbox fields.</param>
    private sealed record AddFormFieldParameters(
        string FormFieldType,
        string FieldName,
        string? DefaultValue,
        string[]? Options,
        bool? CheckedValue);
}
