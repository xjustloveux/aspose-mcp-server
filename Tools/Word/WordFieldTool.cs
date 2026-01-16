using System.ComponentModel;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
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
    ///     Handler registry for field operations
    /// </summary>
    private readonly HandlerRegistry<Document> _handlerRegistry;

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
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Word.Field");
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

        var parameters = BuildParameters(op, fieldType, fieldArgument, paragraphIndex, insertAtStart, fieldIndex,
            fieldCode, lockField, unlockField, updateField, keepResult, updateAll, includeCode, includeResult,
            formFieldType, fieldName, defaultValue, options, checkedValue, value, selectedIndex, fieldNames);

        var handler = _handlerRegistry.GetHandler(op);

        var operationContext = new OperationContext<Document>
        {
            Document = ctx.Document,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = path,
            OutputPath = outputPath
        };

        var result = handler.Execute(operationContext, parameters);

        // Read-only operations don't need to save
        if (op is "get_fields" or "get_field_detail" or "get_form_fields")
            return result;

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return $"{result}\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    /// </summary>
    private static OperationParameters BuildParameters(
        string operation,
        string? fieldType,
        string? fieldArgument,
        int? paragraphIndex,
        bool insertAtStart,
        int? fieldIndex,
        string? fieldCode,
        bool? lockField,
        bool? unlockField,
        bool updateField,
        bool keepResult,
        bool? updateAll,
        bool includeCode,
        bool includeResult,
        string? formFieldType,
        string? fieldName,
        string? defaultValue,
        string[]? options,
        bool? checkedValue,
        string? value,
        int? selectedIndex,
        string[]? fieldNames)
    {
        var parameters = new OperationParameters();

        return operation switch
        {
            "insert_field" => BuildInsertFieldParameters(parameters, fieldType, fieldArgument, paragraphIndex,
                insertAtStart),
            "edit_field" => BuildEditFieldParameters(parameters, fieldIndex, fieldCode, lockField, unlockField,
                updateField),
            "delete_field" => BuildDeleteFieldParameters(parameters, fieldIndex, keepResult),
            "update_field" => BuildUpdateFieldParameters(parameters, fieldIndex, updateAll),
            "get_fields" => BuildGetFieldsParameters(parameters, includeCode, includeResult),
            "get_field_detail" => BuildFieldIndexParameters(parameters, fieldIndex),
            "add_form_field" => BuildAddFormFieldParameters(parameters, formFieldType, fieldName, defaultValue, options,
                checkedValue),
            "edit_form_field" =>
                BuildEditFormFieldParameters(parameters, fieldName, value, checkedValue, selectedIndex),
            "delete_form_field" => BuildDeleteFormFieldParameters(parameters, fieldName, fieldNames),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for the insert field operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="fieldType">The field type (e.g., 'DATE', 'TIME', 'PAGE', 'NUMPAGES', 'AUTHOR').</param>
    /// <param name="fieldArgument">The field argument.</param>
    /// <param name="paragraphIndex">The paragraph index (0-based, -1 for document end).</param>
    /// <param name="insertAtStart">Whether to insert at start of paragraph.</param>
    /// <returns>OperationParameters configured for the insert field operation.</returns>
    private static OperationParameters BuildInsertFieldParameters(OperationParameters parameters, string? fieldType,
        string? fieldArgument, int? paragraphIndex, bool insertAtStart)
    {
        if (fieldType != null) parameters.Set("fieldType", fieldType);
        if (fieldArgument != null) parameters.Set("fieldArgument", fieldArgument);
        if (paragraphIndex.HasValue) parameters.Set("paragraphIndex", paragraphIndex.Value);
        parameters.Set("insertAtStart", insertAtStart);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the edit field operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="fieldIndex">The field index (0-based).</param>
    /// <param name="fieldCode">The new field code.</param>
    /// <param name="lockField">Whether to lock the field.</param>
    /// <param name="unlockField">Whether to unlock the field.</param>
    /// <param name="updateField">Whether to update field after editing.</param>
    /// <returns>OperationParameters configured for the edit field operation.</returns>
    private static OperationParameters BuildEditFieldParameters(OperationParameters parameters, int? fieldIndex,
        string? fieldCode, bool? lockField, bool? unlockField, bool updateField)
    {
        if (fieldIndex.HasValue) parameters.Set("fieldIndex", fieldIndex.Value);
        if (fieldCode != null) parameters.Set("fieldCode", fieldCode);
        if (lockField.HasValue) parameters.Set("lockField", lockField.Value);
        if (unlockField.HasValue) parameters.Set("unlockField", unlockField.Value);
        parameters.Set("updateField", updateField);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the delete field operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="fieldIndex">The field index (0-based).</param>
    /// <param name="keepResult">Whether to keep field result text after deletion.</param>
    /// <returns>OperationParameters configured for the delete field operation.</returns>
    private static OperationParameters BuildDeleteFieldParameters(OperationParameters parameters, int? fieldIndex,
        bool keepResult)
    {
        if (fieldIndex.HasValue) parameters.Set("fieldIndex", fieldIndex.Value);
        parameters.Set("keepResult", keepResult);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the update field operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="fieldIndex">The field index (0-based).</param>
    /// <param name="updateAll">Whether to update all fields.</param>
    /// <returns>OperationParameters configured for the update field operation.</returns>
    private static OperationParameters BuildUpdateFieldParameters(OperationParameters parameters, int? fieldIndex,
        bool? updateAll)
    {
        if (fieldIndex.HasValue) parameters.Set("fieldIndex", fieldIndex.Value);
        if (updateAll.HasValue) parameters.Set("updateAll", updateAll.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the get fields operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="includeCode">Whether to include field code in results.</param>
    /// <param name="includeResult">Whether to include field result in results.</param>
    /// <returns>OperationParameters configured for the get fields operation.</returns>
    private static OperationParameters BuildGetFieldsParameters(OperationParameters parameters, bool includeCode,
        bool includeResult)
    {
        parameters.Set("includeCode", includeCode);
        parameters.Set("includeResult", includeResult);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for field index-based operations (get_field_detail).
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="fieldIndex">The field index (0-based).</param>
    /// <returns>OperationParameters configured for field index-based operations.</returns>
    private static OperationParameters BuildFieldIndexParameters(OperationParameters parameters, int? fieldIndex)
    {
        if (fieldIndex.HasValue) parameters.Set("fieldIndex", fieldIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the add form field operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="formFieldType">The form field type: 'TextInput', 'CheckBox', 'DropDown'.</param>
    /// <param name="fieldName">The field name.</param>
    /// <param name="defaultValue">The default value.</param>
    /// <param name="options">The options for dropdown fields.</param>
    /// <param name="checkedValue">The checked state for checkbox fields.</param>
    /// <returns>OperationParameters configured for the add form field operation.</returns>
    private static OperationParameters BuildAddFormFieldParameters(OperationParameters parameters,
        string? formFieldType, string? fieldName, string? defaultValue, string[]? options, bool? checkedValue)
    {
        if (formFieldType != null) parameters.Set("formFieldType", formFieldType);
        if (fieldName != null) parameters.Set("fieldName", fieldName);
        if (defaultValue != null) parameters.Set("defaultValue", defaultValue);
        if (options != null) parameters.Set("options", options);
        if (checkedValue.HasValue) parameters.Set("checkedValue", checkedValue.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the edit form field operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="fieldName">The field name.</param>
    /// <param name="value">The new value for text input fields.</param>
    /// <param name="checkedValue">The checked state for checkbox fields.</param>
    /// <param name="selectedIndex">The selected option index for dropdown fields.</param>
    /// <returns>OperationParameters configured for the edit form field operation.</returns>
    private static OperationParameters BuildEditFormFieldParameters(OperationParameters parameters, string? fieldName,
        string? value, bool? checkedValue, int? selectedIndex)
    {
        if (fieldName != null) parameters.Set("fieldName", fieldName);
        if (value != null) parameters.Set("value", value);
        if (checkedValue.HasValue) parameters.Set("checkedValue", checkedValue.Value);
        if (selectedIndex.HasValue) parameters.Set("selectedIndex", selectedIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the delete form field operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="fieldName">The field name to delete.</param>
    /// <param name="fieldNames">The array of field names to delete.</param>
    /// <returns>OperationParameters configured for the delete form field operation.</returns>
    private static OperationParameters BuildDeleteFormFieldParameters(OperationParameters parameters, string? fieldName,
        string[]? fieldNames)
    {
        if (fieldName != null) parameters.Set("fieldName", fieldName);
        if (fieldNames != null) parameters.Set("fieldNames", fieldNames);
        return parameters;
    }
}
