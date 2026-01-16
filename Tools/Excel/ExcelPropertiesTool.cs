using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel properties (workbook properties, sheet properties, sheet info).
/// </summary>
[McpServerToolType]
public class ExcelPropertiesTool
{
    /// <summary>
    ///     Handler registry for properties operations.
    /// </summary>
    private readonly HandlerRegistry<Workbook> _handlerRegistry;

    /// <summary>
    ///     Session identity accessor for session isolation support.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Document session manager for in-memory editing support.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelPropertiesTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelPropertiesTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Workbook>.CreateFromNamespace("AsposeMcpServer.Handlers.Excel.Properties");
    }

    /// <summary>
    ///     Executes an Excel properties operation (get_workbook_properties, set_workbook_properties, get_sheet_properties,
    ///     edit_sheet_properties, get_sheet_info).
    /// </summary>
    /// <param name="operation">
    ///     The operation to perform: get_workbook_properties, set_workbook_properties,
    ///     get_sheet_properties, edit_sheet_properties, get_sheet_info.
    /// </param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sheetIndex">Sheet index (0-based, default: 0, required for sheet operations).</param>
    /// <param name="title">Title (optional, for set_workbook_properties).</param>
    /// <param name="subject">Subject (optional, for set_workbook_properties).</param>
    /// <param name="author">Author (optional, for set_workbook_properties).</param>
    /// <param name="keywords">Keywords (optional, for set_workbook_properties).</param>
    /// <param name="comments">Comments (optional, for set_workbook_properties).</param>
    /// <param name="category">Category (optional, for set_workbook_properties).</param>
    /// <param name="company">Company (optional, for set_workbook_properties).</param>
    /// <param name="manager">Manager (optional, for set_workbook_properties).</param>
    /// <param name="customProperties">Custom properties as JSON object (optional, for set_workbook_properties).</param>
    /// <param name="name">Sheet name (optional, for edit_sheet_properties).</param>
    /// <param name="isVisible">Sheet visibility (optional, for edit_sheet_properties).</param>
    /// <param name="tabColor">Tab color hex (e.g., #FF0000, optional, for edit_sheet_properties).</param>
    /// <param name="isSelected">Set as selected sheet (optional, for edit_sheet_properties).</param>
    /// <param name="targetSheetIndex">Sheet index for get_sheet_info (optional, if not provided returns all sheets).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "excel_properties")]
    [Description(
        @"Manage Excel properties. Supports 5 operations: get_workbook_properties, set_workbook_properties, get_sheet_properties, edit_sheet_properties, get_sheet_info.

Usage examples:
- Get workbook properties: excel_properties(operation='get_workbook_properties', path='book.xlsx')
- Set workbook properties: excel_properties(operation='set_workbook_properties', path='book.xlsx', title='Title', author='Author')
- Get sheet properties: excel_properties(operation='get_sheet_properties', path='book.xlsx', sheetIndex=0)
- Edit sheet properties: excel_properties(operation='edit_sheet_properties', path='book.xlsx', sheetIndex=0, name='New Name')
- Get sheet info: excel_properties(operation='get_sheet_info', path='book.xlsx')")]
    public string Execute(
        [Description(@"Operation to perform.
- 'get_workbook_properties': Get workbook properties (required params: path)
- 'set_workbook_properties': Set workbook properties (required params: path)
- 'get_sheet_properties': Get sheet properties (required params: path, sheetIndex)
- 'edit_sheet_properties': Edit sheet properties (required params: path, sheetIndex)
- 'get_sheet_info': Get sheet info (required params: path)")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0, required for sheet operations)")]
        int sheetIndex = 0,
        [Description("Title (optional, for set_workbook_properties)")]
        string? title = null,
        [Description("Subject (optional, for set_workbook_properties)")]
        string? subject = null,
        [Description("Author (optional, for set_workbook_properties)")]
        string? author = null,
        [Description("Keywords (optional, for set_workbook_properties)")]
        string? keywords = null,
        [Description("Comments (optional, for set_workbook_properties)")]
        string? comments = null,
        [Description("Category (optional, for set_workbook_properties)")]
        string? category = null,
        [Description("Company (optional, for set_workbook_properties)")]
        string? company = null,
        [Description("Manager (optional, for set_workbook_properties)")]
        string? manager = null,
        [Description("Custom properties as JSON object (optional, for set_workbook_properties)")]
        string? customProperties = null,
        [Description("Sheet name (optional, for edit_sheet_properties)")]
        string? name = null,
        [Description("Sheet visibility (optional, for edit_sheet_properties)")]
        bool? isVisible = null,
        [Description("Tab color hex (e.g., #FF0000, optional, for edit_sheet_properties)")]
        string? tabColor = null,
        [Description("Set as selected sheet (optional, for edit_sheet_properties)")]
        bool? isSelected = null,
        [Description("Sheet index for get_sheet_info (optional, if not provided returns all sheets)")]
        int? targetSheetIndex = null)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, sheetIndex, title, subject, author, keywords, comments,
            category, company, manager, customProperties, name, isVisible, tabColor, isSelected, targetSheetIndex);

        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<Workbook>
        {
            Document = ctx.Document,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = path,
            OutputPath = outputPath
        };

        var result = handler.Execute(operationContext, parameters);

        var op = operation.ToLowerInvariant();
        if (op == "get_workbook_properties" || op == "get_sheet_properties" || op == "get_sheet_info")
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
        int sheetIndex,
        string? title,
        string? subject,
        string? author,
        string? keywords,
        string? comments,
        string? category,
        string? company,
        string? manager,
        string? customProperties,
        string? name,
        bool? isVisible,
        string? tabColor,
        bool? isSelected,
        int? targetSheetIndex)
    {
        var parameters = new OperationParameters();
        parameters.Set("sheetIndex", sheetIndex);

        return operation.ToLowerInvariant() switch
        {
            "get_workbook_properties" or "get_sheet_properties" => parameters,
            "set_workbook_properties" => BuildSetWorkbookPropertiesParameters(parameters, title, subject, author,
                keywords, comments, category, company, manager, customProperties),
            "edit_sheet_properties" => BuildEditSheetPropertiesParameters(parameters, name, isVisible, tabColor,
                isSelected),
            "get_sheet_info" => BuildGetSheetInfoParameters(parameters, targetSheetIndex),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for the set workbook properties operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="title">The workbook title.</param>
    /// <param name="subject">The workbook subject.</param>
    /// <param name="author">The workbook author.</param>
    /// <param name="keywords">The workbook keywords.</param>
    /// <param name="comments">The workbook comments.</param>
    /// <param name="category">The workbook category.</param>
    /// <param name="company">The workbook company.</param>
    /// <param name="manager">The workbook manager.</param>
    /// <param name="customProperties">Custom properties as JSON string.</param>
    /// <returns>OperationParameters configured for setting workbook properties.</returns>
    private static OperationParameters BuildSetWorkbookPropertiesParameters(OperationParameters parameters,
        string? title, string? subject, string? author, string? keywords, string? comments, string? category,
        string? company, string? manager, string? customProperties)
    {
        if (title != null) parameters.Set("title", title);
        if (subject != null) parameters.Set("subject", subject);
        if (author != null) parameters.Set("author", author);
        if (keywords != null) parameters.Set("keywords", keywords);
        if (comments != null) parameters.Set("comments", comments);
        if (category != null) parameters.Set("category", category);
        if (company != null) parameters.Set("company", company);
        if (manager != null) parameters.Set("manager", manager);
        if (customProperties != null) parameters.Set("customProperties", customProperties);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the edit sheet properties operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="name">The new sheet name.</param>
    /// <param name="isVisible">The sheet visibility status.</param>
    /// <param name="tabColor">The sheet tab color in hex format.</param>
    /// <param name="isSelected">Whether the sheet is selected.</param>
    /// <returns>OperationParameters configured for editing sheet properties.</returns>
    private static OperationParameters BuildEditSheetPropertiesParameters(OperationParameters parameters, string? name,
        bool? isVisible, string? tabColor, bool? isSelected)
    {
        if (name != null) parameters.Set("name", name);
        if (isVisible.HasValue) parameters.Set("isVisible", isVisible.Value);
        if (tabColor != null) parameters.Set("tabColor", tabColor);
        if (isSelected.HasValue) parameters.Set("isSelected", isSelected.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the get sheet info operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="targetSheetIndex">The target sheet index to get info for.</param>
    /// <returns>OperationParameters configured for getting sheet info.</returns>
    private static OperationParameters BuildGetSheetInfoParameters(OperationParameters parameters,
        int? targetSheetIndex)
    {
        if (targetSheetIndex.HasValue) parameters.Set("targetSheetIndex", targetSheetIndex.Value);
        return parameters;
    }
}
