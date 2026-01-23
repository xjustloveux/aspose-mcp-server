using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Word.Field;

namespace AsposeMcpServer.Handlers.Word.Field;

/// <summary>
///     Handler for getting all form fields from Word documents.
/// </summary>
[ResultType(typeof(GetFormFieldsWordResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        _ = parameters;

        var document = context.Document;
        var formFields = document.Range.FormFields.ToList();
        List<FormFieldInfo> formFieldsList = [];

        for (var i = 0; i < formFields.Count; i++)
        {
            var field = formFields[i];
            var fieldData = field.Type switch
            {
                FieldType.FieldFormTextInput => new TextFormFieldInfo
                {
                    Index = i,
                    Name = field.Name,
                    Type = field.Type.ToString(),
                    Value = field.Result
                },
                FieldType.FieldFormCheckBox => new CheckBoxFormFieldInfo
                {
                    Index = i,
                    Name = field.Name,
                    Type = field.Type.ToString(),
                    IsChecked = field.Checked
                },
                FieldType.FieldFormDropDown => new DropDownFormFieldInfo
                {
                    Index = i,
                    Name = field.Name,
                    Type = field.Type.ToString(),
                    SelectedIndex = field.DropDownSelectedIndex,
                    Options = field.DropDownItems.ToList()
                },
                _ => new FormFieldInfo
                {
                    Index = i,
                    Name = field.Name,
                    Type = field.Type.ToString()
                }
            };

            formFieldsList.Add(fieldData);
        }

        return new GetFormFieldsWordResult
        {
            Count = formFields.Count,
            FormFields = formFieldsList
        };
    }
}
