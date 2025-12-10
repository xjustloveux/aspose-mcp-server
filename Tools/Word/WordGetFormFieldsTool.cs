using System.Text.Json.Nodes;
using System.Text;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeMcpServer.Tools;

public class WordGetFormFieldsTool : IAsposeTool
{
    public string Description => "Get all form fields from Word document";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Document file path"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

        var doc = new Document(path);
        var sb = new StringBuilder();

        sb.AppendLine("=== Form Fields ===");
        sb.AppendLine();

        var formFields = doc.Range.FormFields.Cast<FormField>().ToList();
        for (int i = 0; i < formFields.Count; i++)
        {
            var field = formFields[i];
            sb.AppendLine($"[{i + 1}] Name: {field.Name}");
            sb.AppendLine($"    Type: {field.Type}");

            switch (field.Type)
            {
                case FieldType.FieldFormTextInput:
                    sb.AppendLine($"    Value: {field.Result ?? "(empty)"}");
                    break;
                case FieldType.FieldFormCheckBox:
                    sb.AppendLine($"    Checked: {field.Checked}");
                    break;
                case FieldType.FieldFormDropDown:
                    sb.AppendLine($"    Selected Index: {field.DropDownSelectedIndex}");
                    sb.AppendLine($"    Options: {string.Join(", ", field.DropDownItems.Cast<string>())}");
                    break;
            }
            sb.AppendLine();
        }

        sb.AppendLine($"Total Form Fields: {formFields.Count}");

        return await Task.FromResult(sb.ToString());
    }
}

