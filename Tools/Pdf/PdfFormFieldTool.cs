using System.Text;
using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

public class PdfFormFieldTool : IAsposeTool
{
    public string Description => @"Manage form fields in PDF documents. Supports 4 operations: add, delete, edit, get.

Usage examples:
- Add form field: pdf_form_field(operation='add', path='doc.pdf', pageIndex=1, fieldType='TextBox', fieldName='name', x=100, y=100, width=200, height=20)
- Delete form field: pdf_form_field(operation='delete', path='doc.pdf', fieldName='name')
- Edit form field: pdf_form_field(operation='edit', path='doc.pdf', fieldName='name', value='New Value')
- Get form field: pdf_form_field(operation='get', path='doc.pdf', fieldName='name')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add a form field (required params: path, pageIndex, fieldType, fieldName, x, y, width, height)
- 'delete': Delete a form field (required params: path, fieldName)
- 'edit': Edit form field value (required params: path, fieldName)
- 'get': Get form field info (required params: path, fieldName)",
                @enum = new[] { "add", "delete", "edit", "get" }
            },
            path = new
            {
                type = "string",
                description = "PDF file path (required for all operations)"
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
        var operation = ArgumentHelper.GetString(arguments, "operation", "operation");

        return operation.ToLower() switch
        {
            "add" => await AddFormField(arguments),
            "delete" => await DeleteFormField(arguments),
            "edit" => await EditFormField(arguments),
            "get" => await GetFormFields(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    /// Adds a form field to a PDF page
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, pageIndex, fieldType, name, x, y, width, height, optional defaultValue, outputPath</param>
    /// <returns>Success message</returns>
    private async Task<string> AddFormField(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var pageIndex = ArgumentHelper.GetInt(arguments, "pageIndex", "pageIndex");
        var fieldType = ArgumentHelper.GetString(arguments, "fieldType", "fieldType");
        var fieldName = ArgumentHelper.GetString(arguments, "fieldName", "fieldName");
        var x = ArgumentHelper.GetDouble(arguments, "x", "x");
        var y = ArgumentHelper.GetDouble(arguments, "y", "y");
        var width = ArgumentHelper.GetDouble(arguments, "width", "width");
        var height = ArgumentHelper.GetDouble(arguments, "height", "height");
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

    /// <summary>
    /// Deletes a form field from a PDF
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, fieldName, optional outputPath</param>
    /// <returns>Success message</returns>
    private async Task<string> DeleteFormField(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var fieldName = ArgumentHelper.GetString(arguments, "fieldName", "fieldName");

        SecurityHelper.ValidateFilePath(path, "path");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        using var document = new Document(path);
        document.Form.Delete(fieldName);
        document.Save(outputPath);
        return await Task.FromResult($"Successfully deleted form field '{fieldName}'. Output: {outputPath}");
    }

    /// <summary>
    /// Edits a form field in a PDF
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, fieldName, optional value, x, y, width, height, outputPath</param>
    /// <returns>Success message</returns>
    private async Task<string> EditFormField(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var fieldName = ArgumentHelper.GetString(arguments, "fieldName", "fieldName");
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

    /// <summary>
    /// Gets all form fields from the PDF
    /// </summary>
    /// <param name="arguments">JSON arguments (no specific parameters required)</param>
    /// <returns>Formatted string with all form fields</returns>
    private async Task<string> GetFormFields(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);

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

