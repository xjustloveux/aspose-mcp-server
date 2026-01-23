using Aspose.Pdf;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Pdf.FormField;

/// <summary>
///     Handler for adding form fields to PDF documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractAddParameters(parameters);

        var document = context.Document;

        if (p.PageIndex < 1 || p.PageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        if (document.Form.Cast<Field>().Any(f => f.PartialName == p.FieldName))
            throw new ArgumentException($"Form field '{p.FieldName}' already exists");

        var page = document.Pages[p.PageIndex];
        var rect = new Rectangle(p.X, p.Y, p.X + p.Width, p.Y + p.Height);
        Field field;

        switch (p.FieldType.ToLower())
        {
            case "textbox":
            case "textfield":
                field = new TextBoxField(page, rect) { PartialName = p.FieldName };
                if (!string.IsNullOrEmpty(p.DefaultValue))
                    ((TextBoxField)field).Value = p.DefaultValue;
                break;
            case "checkbox":
                field = new CheckboxField(page, rect) { PartialName = p.FieldName };
                break;
            case "radiobutton":
                field = new RadioButtonField(page) { PartialName = p.FieldName };
                var optionName = !string.IsNullOrEmpty(p.DefaultValue) ? p.DefaultValue : "Option1";
                var radioOption = new RadioButtonOptionField(page, rect) { OptionName = optionName };
                ((RadioButtonField)field).Add(radioOption);
                break;
            default:
                throw new ArgumentException($"Unknown field type: {p.FieldType}");
        }

        document.Form.Add(field);
        MarkModified(context);

        return new SuccessResult { Message = $"Added {p.FieldType} field '{p.FieldName}'." };
    }

    /// <summary>
    ///     Extracts add parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static AddParameters ExtractAddParameters(OperationParameters parameters)
    {
        return new AddParameters(
            parameters.GetRequired<int>("pageIndex"),
            parameters.GetRequired<string>("fieldType"),
            parameters.GetRequired<string>("fieldName"),
            parameters.GetRequired<double>("x"),
            parameters.GetRequired<double>("y"),
            parameters.GetRequired<double>("width"),
            parameters.GetRequired<double>("height"),
            parameters.GetOptional<string?>("defaultValue"));
    }

    /// <summary>
    ///     Parameters for adding a form field.
    /// </summary>
    /// <param name="PageIndex">The 1-based page index.</param>
    /// <param name="FieldType">The type of form field (textbox, checkbox, radiobutton).</param>
    /// <param name="FieldName">The name of the form field.</param>
    /// <param name="X">The X coordinate.</param>
    /// <param name="Y">The Y coordinate.</param>
    /// <param name="Width">The width of the field.</param>
    /// <param name="Height">The height of the field.</param>
    /// <param name="DefaultValue">The optional default value.</param>
    private sealed record AddParameters(
        int PageIndex,
        string FieldType,
        string FieldName,
        double X,
        double Y,
        double Width,
        double Height,
        string? DefaultValue);
}
