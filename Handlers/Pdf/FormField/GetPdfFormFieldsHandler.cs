using Aspose.Pdf;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Pdf.FormField;

namespace AsposeMcpServer.Handlers.Pdf.FormField;

/// <summary>
///     Handler for retrieving form fields from PDF documents.
/// </summary>
[ResultType(typeof(GetFormFieldsResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractGetParameters(parameters);

        var document = context.Document;

        if (document.Form.Count == 0)
        {
            var emptyResult = new GetFormFieldsResult
            {
                Count = 0,
                TotalCount = 0,
                Truncated = false,
                Items = Array.Empty<PdfFormFieldInfo>(),
                Message = "No form fields found"
            };
            return emptyResult;
        }

        List<PdfFormFieldInfo> fieldList = [];
        foreach (var field in document.Form.Cast<Field>().Take(p.Limit))
        {
            string? value = null;
            bool? isChecked = null;
            int? selected = null;

            if (field is TextBoxField textBox)
                value = textBox.Value;
            else if (field is CheckboxField checkBox)
                isChecked = checkBox.Checked;
            else if (field is RadioButtonField radioButton)
                selected = radioButton.Selected;

            fieldList.Add(new PdfFormFieldInfo
            {
                Name = field.PartialName ?? string.Empty,
                Type = field.GetType().Name,
                Value = value,
                Checked = isChecked,
                Selected = selected
            });
        }

        var totalCount = document.Form.Count;
        var result = new GetFormFieldsResult
        {
            Count = fieldList.Count,
            TotalCount = totalCount,
            Truncated = totalCount > p.Limit,
            Items = fieldList
        };
        return result;
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
