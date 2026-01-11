using System.ComponentModel;
using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Handlers.Pdf.FormField;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing form fields in PDF documents (add, delete, edit, get)
/// </summary>
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
        _handlerRegistry = PdfFormFieldHandlerRegistry.Create();
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
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, pageIndex, fieldType, fieldName, x, y, width, height,
            defaultValue, value, checkedValue, limit);

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

        if (operation.ToLowerInvariant() == "get")
            return result;

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return $"{result}\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters.
    /// </summary>
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
        int limit)
    {
        var parameters = new OperationParameters();

        switch (operation.ToLowerInvariant())
        {
            case "add":
                if (pageIndex.HasValue) parameters.Set("pageIndex", pageIndex.Value);
                if (fieldType != null) parameters.Set("fieldType", fieldType);
                if (fieldName != null) parameters.Set("fieldName", fieldName);
                if (x.HasValue) parameters.Set("x", x.Value);
                if (y.HasValue) parameters.Set("y", y.Value);
                if (width.HasValue) parameters.Set("width", width.Value);
                if (height.HasValue) parameters.Set("height", height.Value);
                if (defaultValue != null) parameters.Set("defaultValue", defaultValue);
                break;

            case "delete":
                if (fieldName != null) parameters.Set("fieldName", fieldName);
                break;

            case "edit":
                if (fieldName != null) parameters.Set("fieldName", fieldName);
                if (value != null) parameters.Set("value", value);
                if (checkedValue.HasValue) parameters.Set("checkedValue", checkedValue.Value);
                break;

            case "get":
                parameters.Set("limit", limit);
                break;
        }

        return parameters;
    }
}
