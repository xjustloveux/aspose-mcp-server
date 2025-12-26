using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing form fields in PDF documents (add, delete, edit, get)
/// </summary>
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

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);

        // Only get outputPath for operations that modify the document
        string? outputPath = null;
        if (operation.ToLower() != "get")
            outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);

        return operation.ToLower() switch
        {
            "add" => await AddFormField(path, outputPath!, arguments),
            "delete" => await DeleteFormField(path, outputPath!, arguments),
            "edit" => await EditFormField(path, outputPath!, arguments),
            "get" => await GetFormFields(path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a form field to a PDF page
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">
    ///     JSON arguments containing pageIndex, fieldType, fieldName, x, y, width, height, optional
    ///     defaultValue
    /// </param>
    /// <returns>Success message</returns>
    private Task<string> AddFormField(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var pageIndex = ArgumentHelper.GetInt(arguments, "pageIndex");
            var fieldType = ArgumentHelper.GetString(arguments, "fieldType");
            var fieldName = ArgumentHelper.GetString(arguments, "fieldName");
            var x = ArgumentHelper.GetDouble(arguments, "x");
            var y = ArgumentHelper.GetDouble(arguments, "y");
            var width = ArgumentHelper.GetDouble(arguments, "width");
            var height = ArgumentHelper.GetDouble(arguments, "height");
            var defaultValue = ArgumentHelper.GetStringNullable(arguments, "defaultValue");

            using var document = new Document(path);
            if (pageIndex < 1 || pageIndex > document.Pages.Count)
                throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

            var page = document.Pages[pageIndex];
            var rect = new Rectangle(x, y, x + width, y + height);
            Field field;

            switch (fieldType.ToLower())
            {
                case "textbox":
                case "textfield":
                    field = new TextBoxField(page, rect) { PartialName = fieldName };
                    if (!string.IsNullOrEmpty(defaultValue))
                        ((TextBoxField)field).Value = defaultValue;
                    break;
                case "checkbox":
                    field = new CheckboxField(page, rect) { PartialName = fieldName };
                    break;
                case "radiobutton":
                    field = new RadioButtonField(page) { PartialName = fieldName };
                    var radioOption = new RadioButtonOptionField(page, rect);
                    ((RadioButtonField)field).Add(radioOption);
                    break;
                default:
                    throw new ArgumentException($"Unknown field type: {fieldType}");
            }

            document.Form.Add(field);
            document.Save(outputPath);
            return $"Added {fieldType} field '{fieldName}'. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes a form field from a PDF
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing fieldName</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteFormField(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var fieldName = ArgumentHelper.GetString(arguments, "fieldName");

            using var document = new Document(path);
            document.Form.Delete(fieldName);
            document.Save(outputPath);
            return $"Deleted form field '{fieldName}'. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Edits a form field in a PDF
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing fieldName, optional value, checkedValue</param>
    /// <returns>Success message</returns>
    private Task<string> EditFormField(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var fieldName = ArgumentHelper.GetString(arguments, "fieldName");
            var value = ArgumentHelper.GetStringNullable(arguments, "value");
            var checkedValue = ArgumentHelper.GetBoolNullable(arguments, "checkedValue");

            using var document = new Document(path);
            var field = document.Form[fieldName];
            if (field == null)
                throw new ArgumentException($"Form field '{fieldName}' not found");

            if (field is TextBoxField textBox && !string.IsNullOrEmpty(value))
                textBox.Value = value;
            else if (field is CheckboxField checkBox && checkedValue.HasValue)
                checkBox.Checked = checkedValue.Value;

            document.Save(outputPath);
            return $"Edited form field '{fieldName}'. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets all form fields from the PDF
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <returns>JSON string with all form fields</returns>
    private Task<string> GetFormFields(string path)
    {
        return Task.Run(() =>
        {
            using var document = new Document(path);

            if (document.Form.Count == 0)
            {
                var emptyResult = new
                {
                    count = 0,
                    items = Array.Empty<object>(),
                    message = "No form fields found"
                };
                return JsonSerializer.Serialize(emptyResult, new JsonSerializerOptions { WriteIndented = true });
            }

            var fieldList = new List<object>();
            foreach (var field in document.Form.Cast<Field>())
            {
                var fieldInfo = new Dictionary<string, object?>
                {
                    ["name"] = field.PartialName,
                    ["type"] = field.GetType().Name
                };
                if (field is TextBoxField textBox)
                    fieldInfo["value"] = textBox.Value;
                else if (field is CheckboxField checkBox)
                    fieldInfo["checked"] = checkBox.Checked;
                fieldList.Add(fieldInfo);
            }

            var result = new
            {
                count = fieldList.Count,
                items = fieldList
            };
            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }
}