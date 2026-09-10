using Aspose.Pdf;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Pdf.FormField;

/// <summary>
///     Handler for editing form field values in PDF documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractEditParameters(parameters);

        var document = context.Document;
        var field = document.Form.Cast<Field>().FirstOrDefault(f => f.PartialName == p.FieldName);
        if (field == null)
            throw new ArgumentException($"Form field '{p.FieldName}' not found");

        if (string.IsNullOrEmpty(p.Value) && !p.CheckedValue.HasValue)
            throw new ArgumentException("Provide 'value' or 'checked' to edit a form field");

        // Reporting success while changing nothing hid two different problems: a caller who sent
        // no field at all, and a field type this handler cannot write.
        var updated = field switch
        {
            TextBoxField textBox when !string.IsNullOrEmpty(p.Value) => Assign(() => textBox.Value = p.Value),
            CheckboxField checkBox when p.CheckedValue.HasValue =>
                Assign(() => checkBox.Checked = p.CheckedValue.Value),
            RadioButtonField radioButton when !string.IsNullOrEmpty(p.Value) =>
                Assign(() => radioButton.Value = p.Value),
            TextBoxField or CheckboxField or RadioButtonField => false,
            _ => throw new ArgumentException(
                $"Editing a field of type '{field.GetType().Name}' is not supported")
        };

        if (!updated)
            throw new ArgumentException(
                $"The supplied values do not apply to form field '{p.FieldName}' of type '{field.GetType().Name}'");

        MarkModified(context);

        return new SuccessResult { Message = $"Edited form field '{p.FieldName}'." };
    }

    /// <summary>
    ///     Runs a field assignment and reports that a change was made, so the switch above can both
    ///     write the value and record that it applied.
    /// </summary>
    /// <param name="assign">The assignment to perform.</param>
    /// <returns>Always <c>true</c>.</returns>
    private static bool Assign(Action assign)
    {
        assign();
        return true;
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
    private sealed record EditParameters(string FieldName, string? Value, bool? CheckedValue);
}
