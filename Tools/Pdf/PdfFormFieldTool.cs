using System.ComponentModel;
using System.Text.Json;
using Aspose.Pdf;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing form fields in PDF documents (add, delete, edit, get)
/// </summary>
[McpServerToolType]
public class PdfFormFieldTool
{
    /// <summary>
    ///     The session identity accessor for session isolation.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     The document session manager for managing in-memory document sessions.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PdfFormFieldTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public PdfFormFieldTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes a PDF form field operation (add, delete, edit, get).
    /// </summary>
    /// <param name="operation">The operation to perform: add, delete, edit, get.</param>
    /// <param name="path">PDF file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="pageIndex">Page index (1-based, required for add).</param>
    /// <param name="fieldType">Field type: TextBox, CheckBox, RadioButton (required for add).</param>
    /// <param name="fieldName">Field name (required for add, delete, edit).</param>
    /// <param name="x">X position in PDF coordinates (required for add).</param>
    /// <param name="y">Y position in PDF coordinates (required for add).</param>
    /// <param name="width">Width (required for add).</param>
    /// <param name="height">Height (required for add).</param>
    /// <param name="defaultValue">Default value (for add, edit).</param>
    /// <param name="value">Field value (for edit).</param>
    /// <param name="checkedValue">Checked state (for CheckBox, RadioButton).</param>
    /// <param name="limit">Maximum number of fields to return (for get, default: 100).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "pdf_form_field")]
    [Description(@"Manage form fields in PDF documents. Supports 4 operations: add, delete, edit, get.

Usage examples:
- Add form field: pdf_form_field(operation='add', path='doc.pdf', pageIndex=1, fieldType='TextBox', fieldName='name', x=100, y=100, width=200, height=20)
- Delete form field: pdf_form_field(operation='delete', path='doc.pdf', fieldName='name')
- Edit form field: pdf_form_field(operation='edit', path='doc.pdf', fieldName='name', value='New Value')
- Get form fields: pdf_form_field(operation='get', path='doc.pdf')
- Get form fields with limit: pdf_form_field(operation='get', path='doc.pdf', limit=50)")]
    public string Execute(
        [Description(@"Operation to perform.
- 'add': Add a form field (required params: path, pageIndex, fieldType, fieldName, x, y, width, height)
- 'delete': Delete a form field (required params: path, fieldName)
- 'edit': Edit form field value (required params: path, fieldName)
- 'get': Get form field info (required params: path, fieldName)")]
        string operation,
        [Description("PDF file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Page index (1-based, required for add)")]
        int? pageIndex = null,
        [Description("Field type: TextBox, CheckBox, RadioButton (required for add)")]
        string? fieldType = null,
        [Description("Field name (required for add, delete, edit)")]
        string? fieldName = null,
        [Description("X position in PDF coordinates, origin at bottom-left corner (required for add)")]
        double? x = null,
        [Description("Y position in PDF coordinates, origin at bottom-left corner (required for add)")]
        double? y = null,
        [Description("Width (required for add)")]
        double? width = null,
        [Description("Height (required for add)")]
        double? height = null,
        [Description("Default value (for add, edit)")]
        string? defaultValue = null,
        [Description("Field value (for edit)")]
        string? value = null,
        [Description("Checked state (for CheckBox, RadioButton)")]
        bool? checkedValue = null,
        [Description("Maximum number of fields to return (for get, default: 100)")]
        int limit = 100)
    {
        return operation.ToLower() switch
        {
            "add" => AddFormField(sessionId, path, outputPath, pageIndex, fieldType, fieldName, x, y, width, height,
                defaultValue),
            "delete" => DeleteFormField(sessionId, path, outputPath, fieldName),
            "edit" => EditFormField(sessionId, path, outputPath, fieldName, value, checkedValue),
            "get" => GetFormFields(sessionId, path, limit),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a new form field to the PDF document.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="pageIndex">The 1-based page index.</param>
    /// <param name="fieldType">The type of form field (TextBox, CheckBox, RadioButton).</param>
    /// <param name="fieldName">The name of the form field.</param>
    /// <param name="x">The X position in PDF coordinates.</param>
    /// <param name="y">The Y position in PDF coordinates.</param>
    /// <param name="width">The width of the field.</param>
    /// <param name="height">The height of the field.</param>
    /// <param name="defaultValue">Optional default value for the field.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    private string AddFormField(string? sessionId, string? path, string? outputPath, int? pageIndex, string? fieldType,
        string? fieldName, double? x, double? y, double? width, double? height, string? defaultValue)
    {
        if (!pageIndex.HasValue)
            throw new ArgumentException("pageIndex is required for add operation");
        if (string.IsNullOrEmpty(fieldType))
            throw new ArgumentException("fieldType is required for add operation");
        if (string.IsNullOrEmpty(fieldName))
            throw new ArgumentException("fieldName is required for add operation");
        if (!x.HasValue)
            throw new ArgumentException("x is required for add operation");
        if (!y.HasValue)
            throw new ArgumentException("y is required for add operation");
        if (!width.HasValue)
            throw new ArgumentException("width is required for add operation");
        if (!height.HasValue)
            throw new ArgumentException("height is required for add operation");

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);
        var document = ctx.Document;

        if (pageIndex.Value < 1 || pageIndex.Value > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        if (document.Form.Cast<Field>().Any(f => f.PartialName == fieldName))
            throw new ArgumentException($"Form field '{fieldName}' already exists");

        var page = document.Pages[pageIndex.Value];
        var rect = new Rectangle(x.Value, y.Value, x.Value + width.Value, y.Value + height.Value);
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
                var optionName = !string.IsNullOrEmpty(defaultValue) ? defaultValue : "Option1";
                var radioOption = new RadioButtonOptionField(page, rect) { OptionName = optionName };
                ((RadioButtonField)field).Add(radioOption);
                break;
            default:
                throw new ArgumentException($"Unknown field type: {fieldType}");
        }

        document.Form.Add(field);
        ctx.Save(outputPath);
        return $"Added {fieldType} field '{fieldName}'. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Deletes a form field from the PDF document.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="fieldName">The name of the form field to delete.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when the field is not found.</exception>
    private string DeleteFormField(string? sessionId, string? path, string? outputPath, string? fieldName)
    {
        if (string.IsNullOrEmpty(fieldName))
            throw new ArgumentException("fieldName is required for delete operation");

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);
        var document = ctx.Document;

        if (document.Form.Cast<Field>().All(f => f.PartialName != fieldName))
            throw new ArgumentException($"Form field '{fieldName}' not found");

        document.Form.Delete(fieldName);
        ctx.Save(outputPath);
        return $"Deleted form field '{fieldName}'. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Edits the value of an existing form field.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="fieldName">The name of the form field to edit.</param>
    /// <param name="value">The new value for text or radio button fields.</param>
    /// <param name="checkedValue">The checked state for checkbox fields.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when the field is not found.</exception>
    private string EditFormField(string? sessionId, string? path, string? outputPath, string? fieldName, string? value,
        bool? checkedValue)
    {
        if (string.IsNullOrEmpty(fieldName))
            throw new ArgumentException("fieldName is required for edit operation");

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);
        var document = ctx.Document;
        var field = document.Form.Cast<Field>().FirstOrDefault(f => f.PartialName == fieldName);
        if (field == null)
            throw new ArgumentException($"Form field '{fieldName}' not found");

        if (field is TextBoxField textBox && !string.IsNullOrEmpty(value))
            textBox.Value = value;
        else if (field is CheckboxField checkBox && checkedValue.HasValue)
            checkBox.Checked = checkedValue.Value;
        else if (field is RadioButtonField radioButton && !string.IsNullOrEmpty(value))
            radioButton.Value = value;

        ctx.Save(outputPath);
        return $"Edited form field '{fieldName}'. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Retrieves all form fields from the PDF document.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <param name="limit">Maximum number of fields to return.</param>
    /// <returns>A JSON string containing form field information.</returns>
    private string GetFormFields(string? sessionId, string? path, int limit)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);
        var document = ctx.Document;

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

        List<object> fieldList = [];
        foreach (var field in document.Form.Cast<Field>().Take(limit))
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
            else if (field is RadioButtonField radioButton)
                fieldInfo["selected"] = radioButton.Selected;
            fieldList.Add(fieldInfo);
        }

        var totalCount = document.Form.Count;
        var result = new
        {
            count = fieldList.Count,
            totalCount,
            truncated = totalCount > limit,
            items = fieldList
        };
        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }
}