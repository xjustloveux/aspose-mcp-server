using System.Text.Json;
using Aspose.Pdf;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Pdf.FormField;

/// <summary>
///     Handler for retrieving form fields from PDF documents.
/// </summary>
public class GetPdfFormFieldsHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets all form fields from the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: limit (maximum number of fields to return, default: 100)
    /// </param>
    /// <returns>JSON string containing form field information.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractGetParameters(parameters);

        var document = context.Document;

        if (document.Form.Count == 0)
        {
            var emptyResult = new
            {
                count = 0,
                items = Array.Empty<object>(),
                message = "No form fields found"
            };
            return JsonSerializer.Serialize(emptyResult, JsonDefaults.Indented);
        }

        List<object> fieldList = [];
        foreach (var field in document.Form.Cast<Field>().Take(p.Limit))
        {
            var fieldInfo = new Dictionary<string, object?>
            {
                ["name"] = field.PartialName,
                ["type"] = field.GetType().Name
            };
            if (field is TextBoxField textBox)
                fieldInfo["value"] = textBox.Value;
            else if (field is CheckboxField checkBox)
                fieldInfo["checked"] = checkBox.Checked;
            else if (field is RadioButtonField radioButton)
                fieldInfo["selected"] = radioButton.Selected;
            fieldList.Add(fieldInfo);
        }

        var totalCount = document.Form.Count;
        var result = new
        {
            count = fieldList.Count,
            totalCount,
            truncated = totalCount > p.Limit,
            items = fieldList
        };
        return JsonSerializer.Serialize(result, JsonDefaults.Indented);
    }

    /// <summary>
    ///     Extracts get parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static GetParameters ExtractGetParameters(OperationParameters parameters)
    {
        return new GetParameters(
            parameters.GetOptional("limit", 100));
    }

    /// <summary>
    ///     Parameters for getting form fields.
    /// </summary>
    /// <param name="Limit">The maximum number of fields to return.</param>
    private sealed record GetParameters(int Limit);
}
