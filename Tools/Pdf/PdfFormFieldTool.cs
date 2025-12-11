using System.Text;
using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

public class PdfFormFieldTool : IAsposeTool
{
    public string Description => "Manage form fields in PDF documents (add, delete, edit, get)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation: add, delete, edit, get",
                @enum = new[] { "add", "delete", "edit", "get" }
            },
            path = new
            {
                type = "string",
                description = "PDF file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to overwrite input)"
            },
            pageIndex = new
            {
                type = "number",
                description = "Page index (1-based, required for add)"
            },
            fieldType = new
            {
                type = "string",
                description = "Field type: TextBox, CheckBox, RadioButton (required for add)",
                @enum = new[] { "TextBox", "CheckBox", "RadioButton" }
            },
            fieldName = new
            {
                type = "string",
                description = "Field name (required for add, delete, edit)"
            },
            x = new
            {
                type = "number",
                description = "X position (required for add)"
            },
            y = new
            {
                type = "number",
                description = "Y position (required for add)"
            },
            width = new
            {
                type = "number",
                description = "Width (required for add)"
            },
            height = new
            {
                type = "number",
                description = "Height (required for add)"
            },
            defaultValue = new
            {
                type = "string",
                description = "Default value (for add, edit)"
            },
            value = new
            {
                type = "string",
                description = "Field value (for edit)"
            },
            checkedValue = new
            {
                type = "boolean",
                description = "Checked state (for CheckBox, RadioButton)"
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");

        return operation.ToLower() switch
        {
            "add" => await AddFormField(arguments),
            "delete" => await DeleteFormField(arguments),
            "edit" => await EditFormField(arguments),
            "get" => await GetFormFields(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> AddFormField(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var pageIndex = arguments?["pageIndex"]?.GetValue<int>() ?? throw new ArgumentException("pageIndex is required");
        var fieldType = arguments?["fieldType"]?.GetValue<string>() ?? throw new ArgumentException("fieldType is required");
        var fieldName = arguments?["fieldName"]?.GetValue<string>() ?? throw new ArgumentException("fieldName is required");
        var x = arguments?["x"]?.GetValue<double>() ?? throw new ArgumentException("x is required");
        var y = arguments?["y"]?.GetValue<double>() ?? throw new ArgumentException("y is required");
        var width = arguments?["width"]?.GetValue<double>() ?? throw new ArgumentException("width is required");
        var height = arguments?["height"]?.GetValue<double>() ?? throw new ArgumentException("height is required");
        var defaultValue = arguments?["defaultValue"]?.GetValue<string>();

        using var document = new Document(path);
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];
        var rect = new Aspose.Pdf.Rectangle(x, y, x + width, y + height);
        Field field;

        switch (fieldType)
        {
            case "TextBox":
                field = new TextBoxField(page, rect) { PartialName = fieldName };
                if (!string.IsNullOrEmpty(defaultValue))
                    ((TextBoxField)field).Value = defaultValue;
                break;
            case "CheckBox":
                field = new CheckboxField(page, rect) { PartialName = fieldName };
                break;
            case "RadioButton":
                field = new RadioButtonField(page) { PartialName = fieldName };
                var radioOption = new RadioButtonOptionField(page, rect);
                ((RadioButtonField)field).Add(radioOption);
                break;
            default:
                throw new ArgumentException($"Unknown field type: {fieldType}");
        }

        document.Form.Add(field);
        document.Save(outputPath);
        return await Task.FromResult($"Successfully added {fieldType} field '{fieldName}'. Output: {outputPath}");
    }

    private async Task<string> DeleteFormField(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var fieldName = arguments?["fieldName"]?.GetValue<string>() ?? throw new ArgumentException("fieldName is required");

        SecurityHelper.ValidateFilePath(path, "path");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        using var document = new Document(path);
        document.Form.Delete(fieldName);
        document.Save(outputPath);
        return await Task.FromResult($"Successfully deleted form field '{fieldName}'. Output: {outputPath}");
    }

    private async Task<string> EditFormField(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var fieldName = arguments?["fieldName"]?.GetValue<string>() ?? throw new ArgumentException("fieldName is required");
        var value = arguments?["value"]?.GetValue<string>();
        var checkedValue = arguments?["checkedValue"]?.GetValue<bool?>();

        SecurityHelper.ValidateFilePath(path, "path");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        using var document = new Document(path);
        var field = document.Form[fieldName];
        if (field == null)
            throw new ArgumentException($"Form field '{fieldName}' not found");

        if (field is TextBoxField textBox && !string.IsNullOrEmpty(value))
            textBox.Value = value;
        else if (field is CheckboxField checkBox && checkedValue.HasValue)
            checkBox.Checked = checkedValue.Value;

        document.Save(outputPath);
        return await Task.FromResult($"Successfully edited form field '{fieldName}'. Output: {outputPath}");
    }

    private async Task<string> GetFormFields(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

        SecurityHelper.ValidateFilePath(path, "path");

        using var document = new Document(path);
        var sb = new StringBuilder();
        sb.AppendLine("=== PDF Form Fields ===");
        sb.AppendLine();

        if (document.Form.Count == 0)
        {
            sb.AppendLine("No form fields found.");
            return await Task.FromResult(sb.ToString());
        }

        sb.AppendLine($"Total Form Fields: {document.Form.Count}");
        sb.AppendLine();

        foreach (Field field in document.Form)
        {
            sb.AppendLine($"Name: {field.PartialName}");
            sb.AppendLine($"Type: {field.GetType().Name}");
            if (field is TextBoxField textBox)
                sb.AppendLine($"Value: {textBox.Value}");
            else if (field is CheckboxField checkBox)
                sb.AppendLine($"Checked: {checkBox.Checked}");
            sb.AppendLine();
        }

        return await Task.FromResult(sb.ToString());
    }
}

