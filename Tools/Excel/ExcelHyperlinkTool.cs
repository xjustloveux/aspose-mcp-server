using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel hyperlinks (add, edit, delete, get).
/// </summary>
[McpServerToolType]
public class ExcelHyperlinkTool
{
    /// <summary>
    ///     Handler registry for hyperlink operations.
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
    ///     Initializes a new instance of the <see cref="ExcelHyperlinkTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelHyperlinkTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Workbook>.CreateFromNamespace("AsposeMcpServer.Handlers.Excel.Hyperlink");
    }

    /// <summary>
    ///     Executes an Excel hyperlink operation (add, edit, delete, get).
    /// </summary>
    /// <param name="operation">The operation to perform: add, edit, delete, get.</param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sheetIndex">Sheet index (0-based, default: 0).</param>
    /// <param name="cell">Cell reference in A1 notation (e.g., 'A1', 'B2'). Required for add, optional for edit/delete.</param>
    /// <param name="url">URL or file path for the hyperlink (required for add, optional for edit).</param>
    /// <param name="displayText">Display text for the hyperlink (optional for add/edit).</param>
    /// <param name="hyperlinkIndex">Hyperlink index (0-based, alternative to cell for edit/delete).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "excel_hyperlink")]
    [Description(@"Manage Excel hyperlinks. Supports 4 operations: add, edit, delete, get.

Usage examples:
- Add hyperlink: excel_hyperlink(operation='add', path='book.xlsx', cell='A1', url='https://example.com', displayText='Link')
- Edit hyperlink: excel_hyperlink(operation='edit', path='book.xlsx', cell='A1', url='https://newurl.com')
- Delete hyperlink: excel_hyperlink(operation='delete', path='book.xlsx', cell='A1')
- Get hyperlinks: excel_hyperlink(operation='get', path='book.xlsx')")]
    public string Execute(
        [Description("Operation: add, edit, delete, get")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description("Cell reference in A1 notation (e.g., 'A1', 'B2'). Required for add, optional for edit/delete.")]
        string? cell = null,
        [Description("URL or file path for the hyperlink (required for add, optional for edit)")]
        string? url = null,
        [Description("Display text for the hyperlink (optional for add/edit)")]
        string? displayText = null,
        [Description("Hyperlink index (0-based, alternative to cell for edit/delete)")]
        int? hyperlinkIndex = null)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, sheetIndex, cell, url, displayText, hyperlinkIndex);

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

        if (string.Equals(operation, "get", StringComparison.OrdinalIgnoreCase))
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
        string? cell,
        string? url,
        string? displayText,
        int? hyperlinkIndex)
    {
        var parameters = new OperationParameters();
        parameters.Set("sheetIndex", sheetIndex);

        return operation.ToLowerInvariant() switch
        {
            "add" => BuildAddParameters(parameters, cell, url, displayText),
            "edit" => BuildEditParameters(parameters, cell, url, displayText, hyperlinkIndex),
            "delete" => BuildDeleteParameters(parameters, cell, hyperlinkIndex),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for the add hyperlink operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="cell">The cell reference for the hyperlink.</param>
    /// <param name="url">The URL or file path for the hyperlink.</param>
    /// <param name="displayText">The display text for the hyperlink.</param>
    /// <returns>OperationParameters configured for adding hyperlink.</returns>
    private static OperationParameters BuildAddParameters(OperationParameters parameters, string? cell, string? url,
        string? displayText)
    {
        if (cell != null) parameters.Set("cell", cell);
        if (url != null) parameters.Set("url", url);
        if (displayText != null) parameters.Set("displayText", displayText);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the edit hyperlink operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="cell">The cell reference for the hyperlink.</param>
    /// <param name="url">The URL or file path for the hyperlink.</param>
    /// <param name="displayText">The display text for the hyperlink.</param>
    /// <param name="hyperlinkIndex">The hyperlink index as alternative to cell.</param>
    /// <returns>OperationParameters configured for editing hyperlink.</returns>
    private static OperationParameters BuildEditParameters(OperationParameters parameters, string? cell, string? url,
        string? displayText, int? hyperlinkIndex)
    {
        if (cell != null) parameters.Set("cell", cell);
        if (url != null) parameters.Set("url", url);
        if (displayText != null) parameters.Set("displayText", displayText);
        if (hyperlinkIndex.HasValue) parameters.Set("hyperlinkIndex", hyperlinkIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the delete hyperlink operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="cell">The cell reference for the hyperlink.</param>
    /// <param name="hyperlinkIndex">The hyperlink index as alternative to cell.</param>
    /// <returns>OperationParameters configured for deleting hyperlink.</returns>
    private static OperationParameters BuildDeleteParameters(OperationParameters parameters, string? cell,
        int? hyperlinkIndex)
    {
        if (cell != null) parameters.Set("cell", cell);
        if (hyperlinkIndex.HasValue) parameters.Set("hyperlinkIndex", hyperlinkIndex.Value);
        return parameters;
    }
}
