using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeMcpServer.Tools;

public class WordEditFormFieldTool : IAsposeTool
{
    public string Description => "Edit form field value in Word document";

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
                description = "Form field name"
            },
            value = new
            {
                type = "string",
                description = "New value (for TextInput)"
            },
            checkedValue = new
            {
                type = "boolean",
                description = "Checked state (for CheckBox)"
            },
            selectedIndex = new
            {
                type = "number",
                description = "Selected option index (for DropDown)"
            }
        },
        required = new[] { "path", "fieldName" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var fieldName = arguments?["fieldName"]?.GetValue<string>() ?? throw new ArgumentException("fieldName is required");
        var value = arguments?["value"]?.GetValue<string>();
        var checkedValue = arguments?["checkedValue"]?.GetValue<bool?>();
        var selectedIndex = arguments?["selectedIndex"]?.GetValue<int?>();

        var doc = new Document(path);
        var field = doc.Range.FormFields[fieldName];

        if (field == null)
        {
            throw new ArgumentException($"Form field '{fieldName}' not found");
        }

        if (field.Type == FieldType.FieldFormTextInput && value != null)
        {
            field.Result = value;
        }
        else if (field.Type == FieldType.FieldFormCheckBox && checkedValue.HasValue)
        {
            field.Checked = checkedValue.Value;
        }
        else if (field.Type == FieldType.FieldFormDropDown && selectedIndex.HasValue)
        {
            if (selectedIndex.Value >= 0 && selectedIndex.Value < field.DropDownItems.Count)
            {
                field.DropDownSelectedIndex = selectedIndex.Value;
            }
        }

        doc.Save(path);
        return await Task.FromResult($"Form field '{fieldName}' updated: {path}");
    }
}

