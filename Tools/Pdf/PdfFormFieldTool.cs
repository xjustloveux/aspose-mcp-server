using System.ComponentModel;
using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing form fields in PDF documents (add, delete, edit, get, export, import, flatten)
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Pdf.FormField")]
[McpServerToolType]
public class PdfFormFieldTool
{
    /// <summary>
    ///     Handler registry for form field operations.
    /// </summary>
    private readonly HandlerRegistry<Document> _handlerRegistry;

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
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Pdf.FormField");
    }

    /// <summary>
    ///     Executes a PDF form field operation (add, delete, edit, get, export, import, flatten).
    /// </summary>
    /// <param name="operation">The operation to perform: add, delete, edit, get, export, import, flatten.</param>
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
    /// <param name="dataPath">Data file path for export/import (FDF, XFDF, or XML).</param>
    /// <param name="format">Data format for export/import: fdf, xfdf, xml (default: xfdf for export, auto-detect for import).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get/export operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "pdf_form_field",
        Title = "PDF Form Field Operations",
        Destructive = true,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(
        @"Manage form fields in PDF documents. Supports 7 operations: add, delete, edit, get, export, import, flatten.

Usage examples:
- Add form field: pdf_form_field(operation='add', path='doc.pdf', pageIndex=1, fieldType='TextBox', fieldName='name', x=100, y=100, width=200, height=20)
- Delete form field: pdf_form_field(operation='delete', path='doc.pdf', fieldName='name')
- Edit form field: pdf_form_field(operation='edit', path='doc.pdf', fieldName='name', value='New Value')
- Get form fields: pdf_form_field(operation='get', path='doc.pdf')
- Export form data: pdf_form_field(operation='export', path='doc.pdf', dataPath='data.xfdf', format='xfdf')
- Import form data: pdf_form_field(operation='import', path='doc.pdf', dataPath='data.xfdf')
- Flatten form fields: pdf_form_field(operation='flatten', path='doc.pdf')")]
    public object Execute(
        [Description(@"Operation to perform.
- 'add': Add a form field (required params: path, pageIndex, fieldType, fieldName, x, y, width, height)
- 'delete': Delete a form field (required params: path, fieldName)
- 'edit': Edit form field value (required params: path, fieldName)
- 'get': Get form field info (required params: path)
- 'export': Export form data to file (required params: path, dataPath; optional: format)
- 'import': Import form data from file (required params: path, dataPath; optional: format)
- 'flatten': Flatten all form fields to static content (required params: path)")]
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
        int limit = 100,
        [Description("Data file path for export/import (FDF, XFDF, or XML file)")]
        string? dataPath = null,
        [Description(
            "Data format for export/import: fdf, xfdf, xml (default: xfdf for export, auto-detect for import)")]
        string? format = null)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, pageIndex, fieldType, fieldName, x, y, width, height,
            defaultValue, value, checkedValue, limit, dataPath, format);

        var handler = _handlerRegistry.GetHandler(operation);

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

        if (operation.ToLowerInvariant() is "get" or "export")
            return ResultHelper.FinalizeResult((dynamic)result, ctx, outputPath);

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return ResultHelper.FinalizeResult((dynamic)result, ctx, outputPath);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters(
        string operation,
        int? pageIndex,
        string? fieldType,
        string? fieldName,
        double? x,
        double? y,
        double? width,
        double? height,
        string? defaultValue,
        string? value,
        bool? checkedValue,
        int limit,
        string? dataPath,
        string? format)
    {
        return operation.ToLowerInvariant() switch
        {
            "add" => BuildAddParameters(pageIndex, fieldType, fieldName, x, y, width, height, defaultValue),
            "delete" => BuildDeleteParameters(fieldName),
            "edit" => BuildEditParameters(fieldName, value, checkedValue),
            "get" => BuildGetParameters(limit),
            "export" => BuildExportParameters(dataPath, format),
            "import" => BuildImportParameters(dataPath, format),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the add form field operation.
    /// </summary>
    /// <param name="pageIndex">The page index (1-based) to add the field to.</param>
    /// <param name="fieldType">The field type (TextBox, CheckBox, RadioButton).</param>
    /// <param name="fieldName">The field name.</param>
    /// <param name="x">The X position in PDF coordinates.</param>
    /// <param name="y">The Y position in PDF coordinates.</param>
    /// <param name="width">The field width.</param>
    /// <param name="height">The field height.</param>
    /// <param name="defaultValue">The default value of the field.</param>
    /// <returns>OperationParameters configured for adding a form field.</returns>
    private static OperationParameters BuildAddParameters(int? pageIndex, string? fieldType, string? fieldName,
        double? x, double? y, double? width, double? height, string? defaultValue)
    {
        var parameters = new OperationParameters();
        if (pageIndex.HasValue) parameters.Set("pageIndex", pageIndex.Value);
        if (fieldType != null) parameters.Set("fieldType", fieldType);
        if (fieldName != null) parameters.Set("fieldName", fieldName);
        if (x.HasValue) parameters.Set("x", x.Value);
        if (y.HasValue) parameters.Set("y", y.Value);
        if (width.HasValue) parameters.Set("width", width.Value);
        if (height.HasValue) parameters.Set("height", height.Value);
        if (defaultValue != null) parameters.Set("defaultValue", defaultValue);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the delete form field operation.
    /// </summary>
    /// <param name="fieldName">The field name to delete.</param>
    /// <returns>OperationParameters configured for deleting a form field.</returns>
    private static OperationParameters BuildDeleteParameters(string? fieldName)
    {
        var parameters = new OperationParameters();
        if (fieldName != null) parameters.Set("fieldName", fieldName);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the edit form field operation.
    /// </summary>
    /// <param name="fieldName">The field name to edit.</param>
    /// <param name="value">The new field value.</param>
    /// <param name="checkedValue">The checked state for CheckBox or RadioButton.</param>
    /// <returns>OperationParameters configured for editing a form field.</returns>
    private static OperationParameters BuildEditParameters(string? fieldName, string? value, bool? checkedValue)
    {
        var parameters = new OperationParameters();
        if (fieldName != null) parameters.Set("fieldName", fieldName);
        if (value != null) parameters.Set("value", value);
        if (checkedValue.HasValue) parameters.Set("checkedValue", checkedValue.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the get form fields operation.
    /// </summary>
    /// <param name="limit">The maximum number of fields to return.</param>
    /// <returns>OperationParameters configured for getting form fields.</returns>
    private static OperationParameters BuildGetParameters(int limit)
    {
        var parameters = new OperationParameters();
        parameters.Set("limit", limit);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the export form data operation.
    /// </summary>
    /// <param name="dataPath">The output file path for exported data.</param>
    /// <param name="format">The export format (fdf, xfdf, xml).</param>
    /// <returns>OperationParameters configured for exporting form data.</returns>
    private static OperationParameters BuildExportParameters(string? dataPath, string? format)
    {
        var parameters = new OperationParameters();
        if (dataPath != null) parameters.Set("dataPath", dataPath);
        if (format != null) parameters.Set("format", format);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the import form data operation.
    /// </summary>
    /// <param name="dataPath">The input data file path.</param>
    /// <param name="format">The import format (fdf, xfdf, xml), or null for auto-detect.</param>
    /// <returns>OperationParameters configured for importing form data.</returns>
    private static OperationParameters BuildImportParameters(string? dataPath, string? format)
    {
        var parameters = new OperationParameters();
        if (dataPath != null) parameters.Set("dataPath", dataPath);
        if (format != null) parameters.Set("format", format);
        return parameters;
    }
}
