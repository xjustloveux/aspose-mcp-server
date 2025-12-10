using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeMcpServer.Tools;

public class WordDeleteFormFieldTool : IAsposeTool
{
    public string Description => "Delete form field(s) from Word document";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Document file path"
            },
            fieldName = new
            {
                type = "string",
                description = "Form field name to delete (optional, if not provided deletes all form fields)"
            },
            fieldNames = new
            {
                type = "array",
                items = new { type = "string" },
                description = "Array of form field names to delete (optional, overrides fieldName)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var fieldName = arguments?["fieldName"]?.GetValue<string>();
        var fieldNamesArray = arguments?["fieldNames"]?.AsArray();

        var doc = new Document(path);
        var formFields = doc.Range.FormFields;

        List<string> fieldsToDelete;
        if (fieldNamesArray != null && fieldNamesArray.Count > 0)
        {
            fieldsToDelete = fieldNamesArray.Select(f => f?.GetValue<string>()).Where(f => !string.IsNullOrEmpty(f)).Select(f => f!).ToList();
        }
        else if (!string.IsNullOrEmpty(fieldName))
        {
            fieldsToDelete = new List<string> { fieldName };
        }
        else
        {
            // Delete all form fields
            fieldsToDelete = formFields.Cast<FormField>().Select(f => f.Name).ToList();
        }

        int deletedCount = 0;
        foreach (var name in fieldsToDelete)
        {
            var field = formFields[name];
            if (field != null)
            {
                field.Remove();
                deletedCount++;
            }
        }

        doc.Save(path);
        return await Task.FromResult($"Deleted {deletedCount} form field(s): {path}");
    }
}

