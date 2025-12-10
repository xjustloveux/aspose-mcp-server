using System.Text.Json.Nodes;
using System.Text;
using Aspose.Pdf;
using Aspose.Pdf.Forms;

namespace AsposeMcpServer.Tools;

public class PdfGetFormFieldsTool : IAsposeTool
{
    public string Description => "Get all form fields from PDF document";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "PDF file path"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

        using var document = new Document(path);
        var sb = new StringBuilder();

        if (document.Form == null || document.Form.Count == 0)
        {
            return await Task.FromResult("No form fields found in the document.");
        }

        sb.AppendLine($"Form Fields ({document.Form.Count}):");
        sb.AppendLine();

        foreach (Field field in document.Form)
        {
            sb.AppendLine($"Field: {field.PartialName ?? "(unnamed)"}");
            sb.AppendLine($"  Type: {field.GetType().Name}");
            sb.AppendLine($"  Full Name: {field.FullName ?? "(none)"}");

            if (field is TextBoxField textBox)
            {
                sb.AppendLine($"  Value: {textBox.Value ?? "(empty)"}");
            }
            else if (field is CheckboxField checkBox)
            {
                sb.AppendLine($"  Checked: {checkBox.Checked}");
            }
            else if (field is RadioButtonOptionField radioButton)
            {
                // Note: RadioButtonOptionField may not have Selected property
                sb.AppendLine($"  Type: RadioButton");
            }

            sb.AppendLine();
        }

        return await Task.FromResult(sb.ToString());
    }
}

