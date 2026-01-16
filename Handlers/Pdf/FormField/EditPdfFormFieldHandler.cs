using Aspose.Pdf;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.FormField;

/// <summary>
///     Handler for editing form field values in PDF documents.
/// </summary>
public class EditPdfFormFieldHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "edit";

    /// <summary>
    ///     Edits the value of an existing form field.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: fieldName
    ///     Optional: value (for text/radio), checkedValue (for checkbox)
    /// </param>
    /// <returns>Success message with edit details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractEditParameters(parameters);

        var document = context.Document;
        var field = document.Form.Cast<Field>().FirstOrDefault(f => f.PartialName == p.FieldName);
        if (field == null)
            throw new ArgumentException($"Form field '{p.FieldName}' not found");

        if (field is TextBoxField textBox && !string.IsNullOrEmpty(p.Value))
            textBox.Value = p.Value;
        else if (field is CheckboxField checkBox && p.CheckedValue.HasValue)
            checkBox.Checked = p.CheckedValue.Value;
        else if (field is RadioButtonField radioButton && !string.IsNullOrEmpty(p.Value))
            radioButton.Value = p.Value;

        MarkModified(context);

        return Success($"Edited form field '{p.FieldName}'.");
    }

    /// <summary>
    ///     Extracts edit parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static EditParameters ExtractEditParameters(OperationParameters parameters)
    {
        return new EditParameters(
            parameters.GetRequired<string>("fieldName"),
            parameters.GetOptional<string?>("value"),
            parameters.GetOptional<bool?>("checkedValue"));
    }

    /// <summary>
    ///     Parameters for editing a form field.
    /// </summary>
    /// <param name="FieldName">The name of the form field to edit.</param>
    /// <param name="Value">The value for text/radio fields.</param>
    /// <param name="CheckedValue">The checked state for checkbox fields.</param>
    private record EditParameters(string FieldName, string? Value, bool? CheckedValue);
}
