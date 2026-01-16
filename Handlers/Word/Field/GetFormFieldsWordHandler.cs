using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Word.Field;

/// <summary>
///     Handler for getting all form fields from Word documents.
/// </summary>
public class GetFormFieldsWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get_form_fields";

    /// <summary>
    ///     Gets all form fields from the document as JSON.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">No required parameters.</param>
    /// <returns>A JSON string containing the list of form fields.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        _ = ExtractGetFormFieldsParameters(parameters);

        var document = context.Document;
        var formFields = document.Range.FormFields.ToList();
        List<object> formFieldsList = [];

        for (var i = 0; i < formFields.Count; i++)
        {
            var field = formFields[i];
            object fieldData = field.Type switch
            {
                FieldType.FieldFormTextInput => new
                {
                    index = i,
                    name = field.Name,
                    type = field.Type.ToString(),
                    value = field.Result
                },
                FieldType.FieldFormCheckBox => new
                {
                    index = i,
                    name = field.Name,
                    type = field.Type.ToString(),
                    isChecked = field.Checked
                },
                FieldType.FieldFormDropDown => new
                {
                    index = i,
                    name = field.Name,
                    type = field.Type.ToString(),
                    selectedIndex = field.DropDownSelectedIndex,
                    options = field.DropDownItems.ToList()
                },
                _ => new
                {
                    index = i,
                    name = field.Name,
                    type = field.Type.ToString()
                }
            };

            formFieldsList.Add(fieldData);
        }

        var result = new
        {
            count = formFields.Count,
            formFields = formFieldsList
        };

        return JsonSerializer.Serialize(result, JsonDefaults.Indented);
    }

    /// <summary>
    ///     Extracts parameters for the get form fields operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static GetFormFieldsParameters ExtractGetFormFieldsParameters(OperationParameters parameters)
    {
        _ = parameters;
        return new GetFormFieldsParameters();
    }

    /// <summary>
    ///     Parameters for the get form fields operation.
    /// </summary>
    private record GetFormFieldsParameters;
}
