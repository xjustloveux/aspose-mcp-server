using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Forms;

namespace AsposeMcpServer.Tools;

public class PdfEditFormFieldTool : IAsposeTool
{
    public string Description => "Edit form field properties (value, position, size, etc.)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "PDF file path"
            },
            fieldName = new
            {
                type = "string",
                description = "Field name (PartialName or FullName)"
            },
            value = new
            {
                type = "string",
                description = "New value (for TextBox, optional)"
            },
            checkedValue = new
            {
                type = "boolean",
                description = "Checked state (for CheckBox, optional)"
            },
            x = new
            {
                type = "number",
                description = "New X position (optional)"
            },
            y = new
            {
                type = "number",
                description = "New Y position (optional)"
            },
            width = new
            {
                type = "number",
                description = "New width (optional)"
            },
            height = new
            {
                type = "number",
                description = "New height (optional)"
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
        var x = arguments?["x"]?.GetValue<double?>();
        var y = arguments?["y"]?.GetValue<double?>();
        var width = arguments?["width"]?.GetValue<double?>();
        var height = arguments?["height"]?.GetValue<double?>();

        using var document = new Document(path);
        if (document.Form == null || document.Form.Count == 0)
        {
            throw new ArgumentException("No form fields found in the document");
        }

        Field? field = null;
        foreach (Field f in document.Form)
        {
            if (f.PartialName == fieldName || f.FullName == fieldName)
            {
                field = f;
                break;
            }
        }

        if (field == null)
        {
            throw new ArgumentException($"Field '{fieldName}' not found");
        }

        var changes = new List<string>();

        if (field is TextBoxField textBox && value != null)
        {
            textBox.Value = value;
            changes.Add($"Value: {value}");
        }

        if (field is CheckboxField checkBox && checkedValue.HasValue)
        {
            checkBox.Checked = checkedValue.Value;
            changes.Add($"Checked: {checkedValue.Value}");
        }

        if (x.HasValue || y.HasValue || width.HasValue || height.HasValue)
        {
            var rect = field.Rect;
            var newX = x ?? rect.LLX;
            var newY = y ?? rect.LLY;
            var newWidth = width ?? (rect.URX - rect.LLX);
            var newHeight = height ?? (rect.URY - rect.LLY);
            field.Rect = new Rectangle(newX, newY, newX + newWidth, newY + newHeight);
            changes.Add($"Position/Size updated");
        }

        document.Save(path);
        return await Task.FromResult($"Form field '{fieldName}' edited: {string.Join(", ", changes)} - {path}");
    }
}

