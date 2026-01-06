using System.ComponentModel;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Tool for managing fields and form fields in Word documents
/// </summary>
[McpServerToolType]
public class WordFieldTool
{
    /// <summary>
    ///     Identity accessor for session isolation
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document session operations
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the WordFieldTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordFieldTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes a Word field operation (insert_field, edit_field, delete_field, update_field, update_all, get_fields,
    ///     get_field_detail, add_form_field, edit_form_field, delete_form_field, get_form_fields).
    /// </summary>
    /// <param name="operation">
    ///     The operation to perform: insert_field, edit_field, delete_field, update_field, update_all,
    ///     get_fields, get_field_detail, add_form_field, edit_form_field, delete_form_field, get_form_fields.
    /// </param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="fieldType">Field type: DATE, TIME, PAGE, NUMPAGES, AUTHOR, etc. (for insert_field).</param>
    /// <param name="fieldArgument">Field argument (for insert_field).</param>
    /// <param name="paragraphIndex">Paragraph index (0-based, -1 for document end, for insert_field).</param>
    /// <param name="insertAtStart">Insert at start of paragraph (for insert_field, default: false).</param>
    /// <param name="fieldIndex">Field index (0-based, for edit_field/delete_field/update_field/get_field_detail).</param>
    /// <param name="fieldCode">New field code (for edit_field).</param>
    /// <param name="lockField">Lock the field (for edit_field).</param>
    /// <param name="unlockField">Unlock the field (for edit_field).</param>
    /// <param name="updateField">Update field after editing (for edit_field, default: true).</param>
    /// <param name="keepResult">Keep field result text after deletion (for delete_field, default: false).</param>
    /// <param name="updateAll">Update all fields (for update_field, default: false if fieldIndex provided).</param>
    /// <param name="includeCode">Include field code in results (for get_fields, default: true).</param>
    /// <param name="includeResult">Include field result in results (for get_fields, default: true).</param>
    /// <param name="formFieldType">Form field type: TextInput, CheckBox, DropDown (for add_form_field).</param>
    /// <param name="fieldName">Field name (for form field operations).</param>
    /// <param name="defaultValue">Default value (for add_form_field/edit_form_field).</param>
    /// <param name="options">Options for dropdown (for add_form_field with DropDown type).</param>
    /// <param name="checkedValue">Checked state (for CheckBox type).</param>
    /// <param name="value">New value (for TextInput type, for edit_form_field).</param>
    /// <param name="selectedIndex">Selected option index (for DropDown type, for edit_form_field).</param>
    /// <param name="fieldNames">Array of form field names to delete (for delete_form_field).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "word_field")]
    [Description(
        @"Manage fields and form fields in Word documents. Supports 11 operations: insert_field, edit_field, delete_field, update_field, update_all, get_fields, get_field_detail, add_form_field, edit_form_field, delete_form_field, get_form_fields.

Usage examples:
- Insert field: word_field(operation='insert_field', path='doc.docx', fieldType='DATE', paragraphIndex=0)
- Edit field: word_field(operation='edit_field', path='doc.docx', fieldIndex=0, fieldArgument='yyyy-MM-dd')
- Delete field: word_field(operation='delete_field', path='doc.docx', fieldIndex=0)
- Update field: word_field(operation='update_field', path='doc.docx', fieldIndex=0)
- Update all fields: word_field(operation='update_all', path='doc.docx')
- Get fields: word_field(operation='get_fields', path='doc.docx')
- Add form field: word_field(operation='add_form_field', path='doc.docx', formFieldType='TextInput', fieldName='name')")]
    public string Execute(
        [Description(
            "Operation: insert_field, edit_field, delete_field, update_field, update_all, get_fields, get_field_detail, add_form_field, edit_form_field, delete_form_field, get_form_fields")]
        string operation,
        [Description("Word document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Field type: DATE, TIME, PAGE, NUMPAGES, AUTHOR, etc. (for insert_field)")]
        string? fieldType = null,
        [Description("Field argument (for insert_field)")]
        string? fieldArgument = null,
        [Description("Paragraph index (0-based, -1 for document end, for insert_field)")]
        int? paragraphIndex = null,
        [Description("Insert at start of paragraph (for insert_field, default: false)")]
        bool insertAtStart = false,
        [Description("Field index (0-based, for edit_field/delete_field/update_field/get_field_detail)")]
        int? fieldIndex = null,
        [Description("New field code (for edit_field)")]
        string? fieldCode = null,
        [Description("Lock the field (for edit_field)")]
        bool? lockField = null,
        [Description("Unlock the field (for edit_field)")]
        bool? unlockField = null,
        [Description("Update field after editing (for edit_field, default: true)")]
        bool updateField = true,
        [Description("Keep field result text after deletion (for delete_field, default: false)")]
        bool keepResult = false,
        [Description("Update all fields (for update_field, default: false if fieldIndex provided)")]
        bool? updateAll = null,
        [Description("Include field code in results (for get_fields, default: true)")]
        bool includeCode = true,
        [Description("Include field result in results (for get_fields, default: true)")]
        bool includeResult = true,
        [Description("Form field type: TextInput, CheckBox, DropDown (for add_form_field)")]
        string? formFieldType = null,
        [Description("Field name (for form field operations)")]
        string? fieldName = null,
        [Description("Default value (for add_form_field/edit_form_field)")]
        string? defaultValue = null,
        [Description("Options for dropdown (for add_form_field with DropDown type)")]
        string[]? options = null,
        [Description("Checked state (for CheckBox type)")]
        bool? checkedValue = null,
        [Description("New value (for TextInput type, for edit_form_field)")]
        string? value = null,
        [Description("Selected option index (for DropDown type, for edit_form_field)")]
        int? selectedIndex = null,
        [Description("Array of form field names to delete (for delete_form_field)")]
        string[]? fieldNames = null)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var op = operation.ToLower();
        if (op == "update_all")
        {
            op = "update_field";
            updateAll = true;
        }

        return op switch
        {
            "insert_field" => InsertField(ctx, outputPath, fieldType, fieldArgument, paragraphIndex, insertAtStart),
            "edit_field" => EditField(ctx, outputPath, fieldIndex, fieldCode, lockField, unlockField, updateField),
            "delete_field" => DeleteField(ctx, outputPath, fieldIndex, keepResult),
            "update_field" => UpdateField(ctx, outputPath, fieldIndex, updateAll ?? !fieldIndex.HasValue),
            "get_fields" => GetFields(ctx, includeCode, includeResult),
            "get_field_detail" => GetFieldDetail(ctx, fieldIndex),
            "add_form_field" => AddFormField(ctx, outputPath, formFieldType, fieldName, defaultValue, options,
                checkedValue),
            "edit_form_field" => EditFormField(ctx, outputPath, fieldName, value, checkedValue, selectedIndex),
            "delete_form_field" => DeleteFormField(ctx, outputPath, fieldName, fieldNames),
            "get_form_fields" => GetFormFields(ctx),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Inserts a field at the specified paragraph position.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="fieldType">The field type (DATE, TIME, PAGE, NUMPAGES, AUTHOR, etc.).</param>
    /// <param name="fieldArgument">The field argument.</param>
    /// <param name="paragraphIndex">The zero-based paragraph index, or -1 for document end.</param>
    /// <param name="insertAtStart">Whether to insert at the start of the paragraph.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when fieldType is empty or paragraph index is out of range.</exception>
    /// <exception cref="InvalidOperationException">Thrown when unable to find the target paragraph.</exception>
    private static string InsertField(DocumentContext<Document> ctx, string? outputPath, string? fieldType,
        string? fieldArgument, int? paragraphIndex, bool insertAtStart)
    {
        if (string.IsNullOrEmpty(fieldType))
            throw new ArgumentException("fieldType is required for insert_field operation");

        var document = ctx.Document;
        var builder = new DocumentBuilder(document);

        if (paragraphIndex.HasValue)
        {
            var paragraphs = document.GetChildNodes(NodeType.Paragraph, true);
            if (paragraphIndex.Value == -1)
            {
                builder.MoveToDocumentEnd();
            }
            else if (paragraphIndex.Value >= 0 && paragraphIndex.Value < paragraphs.Count)
            {
                if (paragraphs[paragraphIndex.Value] is Paragraph targetPara)
                {
                    if (insertAtStart)
                    {
                        builder.MoveTo(targetPara);
                        if (targetPara.Runs.Count > 0)
                            builder.MoveTo(targetPara.Runs[0]);
                    }
                    else
                    {
                        builder.MoveTo(targetPara);
                        if (targetPara.Runs.Count > 0)
                            builder.MoveTo(targetPara.Runs[^1]);
                    }
                }
                else
                {
                    throw new InvalidOperationException($"Unable to find paragraph at index {paragraphIndex.Value}");
                }
            }
            else
            {
                throw new ArgumentException(
                    $"Paragraph index {paragraphIndex.Value} is out of range (document has {paragraphs.Count} paragraphs)");
            }
        }
        else
        {
            builder.MoveToDocumentEnd();
        }

        var code = fieldType.ToUpper();
        if (!string.IsNullOrEmpty(fieldArgument))
            code += " " + fieldArgument;

        var field = builder.InsertField(code);
        field.Update();

        ctx.Save(outputPath);

        var result = $"Field inserted successfully\nField type: {fieldType}\n";
        if (!string.IsNullOrEmpty(fieldArgument))
            result += $"Field argument: {fieldArgument}\n";
        result += $"Field code: {code}\n";

        try
        {
            var fieldResult = field.Result;
            if (!string.IsNullOrEmpty(fieldResult))
                result += $"Field result: {fieldResult}\n";
        }
        catch
        {
            // Ignore errors reading field result (some fields may not have results)
        }

        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Edits a field's code, lock state, or triggers an update.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="fieldIndex">The zero-based field index.</param>
    /// <param name="fieldCode">The new field code.</param>
    /// <param name="lockField">Whether to lock the field.</param>
    /// <param name="unlockField">Whether to unlock the field.</param>
    /// <param name="updateFieldAfter">Whether to update the field after editing.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when fieldIndex is not provided or is out of range.</exception>
    private static string EditField(DocumentContext<Document> ctx, string? outputPath, int? fieldIndex,
        string? fieldCode, bool? lockField, bool? unlockField, bool updateFieldAfter)
    {
        if (!fieldIndex.HasValue)
            throw new ArgumentException("fieldIndex is required for edit_field operation");

        var document = ctx.Document;
        var fields = document.Range.Fields.ToList();

        if (fieldIndex.Value < 0 || fieldIndex.Value >= fields.Count)
            throw new ArgumentException(
                $"Field index {fieldIndex.Value} is out of range (document has {fields.Count} fields)");

        var field = fields[fieldIndex.Value];
        var oldFieldCode = field.GetFieldCode();
        List<string> changes = [];

        if (!string.IsNullOrEmpty(fieldCode))
        {
            var fieldStart = field.Start;
            var fieldEnd = field.End;

            if (fieldStart != null && fieldEnd != null)
            {
                var builder = new DocumentBuilder(document);
                builder.MoveTo(fieldStart);

                var currentNode = fieldStart.NextSibling;
                while (currentNode != null && currentNode != fieldEnd)
                {
                    var nextNode = currentNode.NextSibling;
                    if (currentNode.NodeType != NodeType.FieldSeparator && currentNode.NodeType != NodeType.FieldEnd)
                        currentNode.Remove();
                    currentNode = nextNode;
                }

                builder.MoveTo(fieldStart);
                builder.Write(fieldCode);
                changes.Add($"Field code updated: {oldFieldCode} -> {fieldCode}");
            }
        }

        if (lockField == true)
        {
            field.IsLocked = true;
            changes.Add("Field locked");
        }
        else if (unlockField == true)
        {
            field.IsLocked = false;
            changes.Add("Field unlocked");
        }

        if (updateFieldAfter)
        {
            field.Update();
            document.UpdateFields();
        }

        ctx.Save(outputPath);

        var result = $"Field #{fieldIndex.Value} edited successfully\n";
        result += $"Original field code: {oldFieldCode}\n";
        if (changes.Count > 0)
            result += $"Changes: {string.Join(", ", changes)}\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Deletes a field from the document, optionally keeping its result text.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="fieldIndex">The zero-based field index.</param>
    /// <param name="keepResult">Whether to keep the field result text after deletion.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when fieldIndex is not provided or is out of range.</exception>
    private static string DeleteField(DocumentContext<Document> ctx, string? outputPath, int? fieldIndex,
        bool keepResult)
    {
        if (!fieldIndex.HasValue)
            throw new ArgumentException("fieldIndex is required for delete_field operation");

        var document = ctx.Document;
        var fields = document.Range.Fields.ToList();

        if (fieldIndex.Value < 0 || fieldIndex.Value >= fields.Count)
            throw new ArgumentException(
                $"Field index {fieldIndex.Value} is out of range (document has {fields.Count} fields)");

        var field = fields[fieldIndex.Value];
        var fieldType = field.Type.ToString();
        var fieldCodeStr = field.GetFieldCode();

        if (keepResult)
            field.Unlink();
        else
            field.Remove();

        ctx.Save(outputPath);

        var remainingFields = document.Range.Fields.Count;
        var result = $"Field #{fieldIndex.Value} deleted successfully\n";
        result += $"Type: {fieldType}\nCode: {fieldCodeStr}\n";
        result += $"Keep result text: {(keepResult ? "Yes" : "No")}\n";
        result += $"Remaining fields: {remainingFields}\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Updates one or all fields in the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="fieldIndex">The zero-based field index.</param>
    /// <param name="updateAllFields">Whether to update all fields in the document.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when fieldIndex is out of range.</exception>
    private static string UpdateField(DocumentContext<Document> ctx, string? outputPath, int? fieldIndex,
        bool updateAllFields)
    {
        var document = ctx.Document;
        var fields = document.Range.Fields.ToList();

        if (fieldIndex.HasValue && !updateAllFields)
        {
            if (fieldIndex.Value < 0 || fieldIndex.Value >= fields.Count)
                throw new ArgumentException(
                    $"Field index {fieldIndex.Value} is out of range (document has {fields.Count} fields)");

            var field = fields[fieldIndex.Value];
            if (field.IsLocked)
                return $"Warning: Field #{fieldIndex.Value} is locked and cannot be updated.";

            var oldResult = field.Result ?? "";
            field.Update();
            var newResult = field.Result ?? "";

            ctx.Save(outputPath);

            return
                $"Field #{fieldIndex.Value} updated\nOld result: {oldResult}\nNew result: {newResult}\n{ctx.GetOutputMessage(outputPath)}";
        }

        var lockedCount = fields.Count(f => f.IsLocked);
        document.UpdateFields();
        ctx.Save(outputPath);

        var result = $"Updated {fields.Count - lockedCount} field(s)\n";
        if (lockedCount > 0)
            result += $"Skipped {lockedCount} locked field(s)\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Gets all fields from the document as JSON with statistics.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="includeCode">Whether to include field code in results.</param>
    /// <param name="includeResult">Whether to include field result in results.</param>
    /// <returns>A JSON string containing the list of fields and statistics.</returns>
    private static string GetFields(DocumentContext<Document> ctx, bool includeCode, bool includeResult)
    {
        var document = ctx.Document;
        List<object> fieldsList = [];
        var fieldIndex = 0;

        foreach (var field in document.Range.Fields)
        {
            string? extraInfo = null;
            if (field is FieldHyperlink hyperlinkField)
                extraInfo = $"Address: {hyperlinkField.Address ?? ""}, ScreenTip: {hyperlinkField.ScreenTip ?? ""}";
            else if (field is FieldRef refField)
                extraInfo = $"Bookmark: {refField.BookmarkName ?? ""}";

            fieldsList.Add(new
            {
                index = fieldIndex++,
                type = field.Type.ToString(),
                code = includeCode ? field.GetFieldCode() : null,
                result = includeResult ? field.Result ?? "" : null,
                isLocked = field.IsLocked,
                isDirty = field.IsDirty,
                extraInfo
            });
        }

        var statistics = fieldsList
            .GroupBy(f => ((dynamic)f).type as string)
            .OrderBy(g => g.Key)
            .Select(g => new { type = g.Key, count = g.Count() })
            .ToList();

        var result = new
        {
            count = fieldsList.Count,
            fields = fieldsList,
            statisticsByType = statistics
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Gets detailed information about a specific field.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="fieldIndex">The zero-based field index.</param>
    /// <returns>A JSON string containing the field details.</returns>
    /// <exception cref="ArgumentException">Thrown when fieldIndex is not provided or is out of range.</exception>
    private static string GetFieldDetail(DocumentContext<Document> ctx, int? fieldIndex)
    {
        if (!fieldIndex.HasValue)
            throw new ArgumentException("fieldIndex is required for get_field_detail operation");

        var document = ctx.Document;
        var fields = document.Range.Fields.ToList();

        if (fieldIndex.Value < 0 || fieldIndex.Value >= fields.Count)
            throw new ArgumentException(
                $"Field index {fieldIndex.Value} is out of range (document has {fields.Count} fields)");

        var field = fields[fieldIndex.Value];

        string? address = null, screenTip = null, bookmarkName = null;
        if (field is FieldHyperlink hyperlinkField)
        {
            address = hyperlinkField.Address;
            screenTip = hyperlinkField.ScreenTip;
        }
        else if (field is FieldRef refField)
        {
            bookmarkName = refField.BookmarkName;
        }

        var result = new
        {
            index = fieldIndex.Value,
            type = field.Type.ToString(),
            typeCode = (int)field.Type,
            code = field.GetFieldCode(),
            result = field.Result,
            isLocked = field.IsLocked,
            isDirty = field.IsDirty,
            hyperlinkAddress = address,
            hyperlinkScreenTip = screenTip,
            bookmarkName
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Adds a form field (text input, checkbox, or dropdown) to the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="formFieldType">The form field type (TextInput, CheckBox, DropDown).</param>
    /// <param name="fieldName">The field name.</param>
    /// <param name="defaultValue">The default value for text input.</param>
    /// <param name="options">The options for dropdown.</param>
    /// <param name="checkedValue">The checked state for checkbox.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when formFieldType or fieldName is empty, or when options is missing for
    ///     dropdown.
    /// </exception>
    private static string AddFormField(DocumentContext<Document> ctx, string? outputPath, string? formFieldType,
        string? fieldName, string? defaultValue, string[]? options, bool? checkedValue)
    {
        if (string.IsNullOrEmpty(formFieldType))
            throw new ArgumentException("formFieldType is required for add_form_field operation");
        if (string.IsNullOrEmpty(fieldName))
            throw new ArgumentException("fieldName is required for add_form_field operation");

        var document = ctx.Document;
        var builder = new DocumentBuilder(document);
        builder.MoveToDocumentEnd();

        switch (formFieldType.ToLower())
        {
            case "textinput":
                builder.InsertTextInput(fieldName, TextFormFieldType.Regular, "", defaultValue ?? "", 0);
                break;
            case "checkbox":
                builder.InsertCheckBox(fieldName, checkedValue ?? false, 0);
                break;
            case "dropdown":
                if (options == null || options.Length == 0)
                    throw new ArgumentException("options array is required for DropDown type");
                builder.InsertComboBox(fieldName, options, 0);
                break;
            default:
                throw new ArgumentException(
                    $"Invalid formFieldType: {formFieldType}. Must be 'TextInput', 'CheckBox', or 'DropDown'");
        }

        ctx.Save(outputPath);
        return $"{formFieldType} field '{fieldName}' added. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Edits an existing form field's value or state.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="fieldName">The field name.</param>
    /// <param name="value">The new value for text input.</param>
    /// <param name="checkedValue">The new checked state for checkbox.</param>
    /// <param name="selectedIndex">The selected option index for dropdown.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when fieldName is empty or the form field is not found.</exception>
    private static string EditFormField(DocumentContext<Document> ctx, string? outputPath, string? fieldName,
        string? value, bool? checkedValue, int? selectedIndex)
    {
        if (string.IsNullOrEmpty(fieldName))
            throw new ArgumentException("fieldName is required for edit_form_field operation");

        var document = ctx.Document;
        var field = document.Range.FormFields[fieldName];

        if (field == null)
            throw new ArgumentException($"Form field '{fieldName}' not found");

        if (field.Type == FieldType.FieldFormTextInput && value != null)
            field.Result = value;
        else if (field.Type == FieldType.FieldFormCheckBox && checkedValue.HasValue)
            field.Checked = checkedValue.Value;
        else if (field.Type == FieldType.FieldFormDropDown && selectedIndex.HasValue)
            if (selectedIndex.Value >= 0 && selectedIndex.Value < field.DropDownItems.Count)
                field.DropDownSelectedIndex = selectedIndex.Value;

        ctx.Save(outputPath);
        return $"Form field '{fieldName}' updated. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Deletes one or more form fields from the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="fieldName">The field name to delete.</param>
    /// <param name="fieldNames">The array of field names to delete.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string DeleteFormField(DocumentContext<Document> ctx, string? outputPath, string? fieldName,
        string[]? fieldNames)
    {
        var document = ctx.Document;
        var formFields = document.Range.FormFields;

        List<string> fieldsToDelete;
        if (fieldNames is { Length: > 0 })
            fieldsToDelete = fieldNames.Where(f => !string.IsNullOrEmpty(f)).ToList();
        else if (!string.IsNullOrEmpty(fieldName))
            fieldsToDelete = [fieldName];
        else
            fieldsToDelete = formFields.Select(f => f.Name).ToList();

        var deletedCount = 0;
        foreach (var name in fieldsToDelete)
        {
            var field = formFields[name];
            if (field != null)
            {
                field.Remove();
                deletedCount++;
            }
        }

        ctx.Save(outputPath);
        return $"Deleted {deletedCount} form field(s). {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Gets all form fields from the document as JSON.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <returns>A JSON string containing the list of form fields.</returns>
    private static string GetFormFields(DocumentContext<Document> ctx)
    {
        var document = ctx.Document;
        var formFields = document.Range.FormFields.ToList();
        List<object> formFieldsList = [];

        for (var i = 0; i < formFields.Count; i++)
        {
            var field = formFields[i];
            object fieldData = field.Type switch
            {
                FieldType.FieldFormTextInput => new
                {
                    index = i,
                    name = field.Name,
                    type = field.Type.ToString(),
                    value = field.Result
                },
                FieldType.FieldFormCheckBox => new
                {
                    index = i,
                    name = field.Name,
                    type = field.Type.ToString(),
                    isChecked = field.Checked
                },
                FieldType.FieldFormDropDown => new
                {
                    index = i,
                    name = field.Name,
                    type = field.Type.ToString(),
                    selectedIndex = field.DropDownSelectedIndex,
                    options = field.DropDownItems.ToList()
                },
                _ => new
                {
                    index = i,
                    name = field.Name,
                    type = field.Type.ToString()
                }
            };

            formFieldsList.Add(fieldData);
        }

        var result = new
        {
            count = formFields.Count,
            formFields = formFieldsList
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }
}