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
        var fieldName = parameters.GetRequired<string>("fieldName");
        var value = parameters.GetOptional<string?>("value");
        var checkedValue = parameters.GetOptional<bool?>("checkedValue");

        var document = context.Document;
        var field = document.Form.Cast<Field>().FirstOrDefault(f => f.PartialName == fieldName);
        if (field == null)
            throw new ArgumentException($"Form field '{fieldName}' not found");

        if (field is TextBoxField textBox && !string.IsNullOrEmpty(value))
            textBox.Value = value;
        else if (field is CheckboxField checkBox && checkedValue.HasValue)
            checkBox.Checked = checkedValue.Value;
        else if (field is RadioButtonField radioButton && !string.IsNullOrEmpty(value))
            radioButton.Value = value;

        MarkModified(context);

        return Success($"Edited form field '{fieldName}'.");
    }
}
