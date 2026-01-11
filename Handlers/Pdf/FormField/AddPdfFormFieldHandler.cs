using Aspose.Pdf;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.FormField;

/// <summary>
///     Handler for adding form fields to PDF documents.
/// </summary>
public class AddPdfFormFieldHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds a new form field to the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: pageIndex (1-based), fieldType, fieldName, x, y, width, height
    ///     Optional: defaultValue
    /// </param>
    /// <returns>Success message with field creation details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var pageIndex = parameters.GetRequired<int>("pageIndex");
        var fieldType = parameters.GetRequired<string>("fieldType");
        var fieldName = parameters.GetRequired<string>("fieldName");
        var x = parameters.GetRequired<double>("x");
        var y = parameters.GetRequired<double>("y");
        var width = parameters.GetRequired<double>("width");
        var height = parameters.GetRequired<double>("height");
        var defaultValue = parameters.GetOptional<string?>("defaultValue");

        var document = context.Document;

        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        if (document.Form.Cast<Field>().Any(f => f.PartialName == fieldName))
            throw new ArgumentException($"Form field '{fieldName}' already exists");

        var page = document.Pages[pageIndex];
        var rect = new Rectangle(x, y, x + width, y + height);
        Field field;

        switch (fieldType.ToLower())
        {
            case "textbox":
            case "textfield":
                field = new TextBoxField(page, rect) { PartialName = fieldName };
                if (!string.IsNullOrEmpty(defaultValue))
                    ((TextBoxField)field).Value = defaultValue;
                break;
            case "checkbox":
                field = new CheckboxField(page, rect) { PartialName = fieldName };
                break;
            case "radiobutton":
                field = new RadioButtonField(page) { PartialName = fieldName };
                var optionName = !string.IsNullOrEmpty(defaultValue) ? defaultValue : "Option1";
                var radioOption = new RadioButtonOptionField(page, rect) { OptionName = optionName };
                ((RadioButtonField)field).Add(radioOption);
                break;
            default:
                throw new ArgumentException($"Unknown field type: {fieldType}");
        }

        document.Form.Add(field);
        MarkModified(context);

        return Success($"Added {fieldType} field '{fieldName}'.");
    }
}
