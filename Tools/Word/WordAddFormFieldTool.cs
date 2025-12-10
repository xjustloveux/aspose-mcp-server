using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeMcpServer.Tools;

public class WordAddFormFieldTool : IAsposeTool
{
    public string Description => "Add form field (text input, checkbox, dropdown) to Word document";

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
            fieldType = new
            {
                type = "string",
                description = "Field type: 'TextInput', 'CheckBox', 'DropDown'",
                @enum = new[] { "TextInput", "CheckBox", "DropDown" }
            },
            fieldName = new
            {
                type = "string",
                description = "Field name"
            },
            defaultValue = new
            {
                type = "string",
                description = "Default value (optional, for TextInput)"
            },
            options = new
            {
                type = "array",
                items = new { type = "string" },
                description = "Options for dropdown (required for DropDown type)"
            },
            checkedValue = new
            {
                type = "boolean",
                description = "Checked state (optional, for CheckBox)"
            }
        },
        required = new[] { "path", "fieldType", "fieldName" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var fieldType = arguments?["fieldType"]?.GetValue<string>() ?? throw new ArgumentException("fieldType is required");
        var fieldName = arguments?["fieldName"]?.GetValue<string>() ?? throw new ArgumentException("fieldName is required");
        var defaultValue = arguments?["defaultValue"]?.GetValue<string>();
        var optionsArray = arguments?["options"]?.AsArray();
        var checkedValue = arguments?["checkedValue"]?.GetValue<bool?>();

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();

        FormField field;
        switch (fieldType.ToLower())
        {
            case "textinput":
                field = builder.InsertTextInput(fieldName, TextFormFieldType.Regular, "", defaultValue ?? "", 0);
                break;

            case "checkbox":
                field = builder.InsertCheckBox(fieldName, checkedValue ?? false, 0);
                break;

            case "dropdown":
                if (optionsArray == null || optionsArray.Count == 0)
                {
                    throw new ArgumentException("options array is required for DropDown type");
                }
                var options = optionsArray.Select(o => o?.GetValue<string>()).Where(o => !string.IsNullOrEmpty(o)).ToArray();
                field = builder.InsertComboBox(fieldName, options, 0);
                break;

            default:
                throw new ArgumentException($"Invalid fieldType: {fieldType}. Must be 'TextInput', 'CheckBox', or 'DropDown'");
        }

        doc.Save(path);
        return await Task.FromResult($"{fieldType} field '{fieldName}' added: {path}");
    }
}

