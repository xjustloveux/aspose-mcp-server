using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel protection (protect, unprotect, get, set_cell_locked).
/// </summary>
[McpServerToolType]
public class ExcelProtectTool
{
    /// <summary>
    ///     Handler registry for protection operations.
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
    ///     Initializes a new instance of the <see cref="ExcelProtectTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelProtectTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Workbook>.CreateFromNamespace("AsposeMcpServer.Handlers.Excel.Protect");
    }

    /// <summary>
    ///     Executes an Excel protection operation (protect, unprotect, get, set_cell_locked).
    /// </summary>
    /// <param name="operation">The operation to perform: protect, unprotect, get, set_cell_locked.</param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sheetIndex">
    ///     Sheet index (0-based, optional). If not specified, the operation applies to the entire
    ///     workbook structure.
    /// </param>
    /// <param name="password">Protection password (required for protect, optional for unprotect).</param>
    /// <param name="protectWorkbook">Protect workbook structure (optional, for protect operation, default: false).</param>
    /// <param name="protectStructure">
    ///     Protect workbook structure (optional, for protect operation when protectWorkbook is
    ///     true, default: true).
    /// </param>
    /// <param name="protectWindows">
    ///     Protect workbook windows (optional, for protect operation when protectWorkbook is true,
    ///     default: false).
    /// </param>
    /// <param name="range">Cell or range (e.g., 'A1' or 'A1:C5', required for set_cell_locked).</param>
    /// <param name="locked">Locked status (true = locked, false = unlocked, required for set_cell_locked).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "excel_protect")]
    [Description(@"Manage Excel protection. Supports 4 operations: protect, unprotect, get, set_cell_locked.

Usage examples:
- Protect sheet: excel_protect(operation='protect', path='book.xlsx', sheetIndex=0, password='password')
- Unprotect sheet: excel_protect(operation='unprotect', path='book.xlsx', sheetIndex=0, password='password')
- Get protection: excel_protect(operation='get', path='book.xlsx', sheetIndex=0)
- Set cell locked: excel_protect(operation='set_cell_locked', path='book.xlsx', range='A1:B10', locked=true)")]
    public string Execute(
        [Description(@"Operation to perform.
- 'protect': Protect workbook or sheet (required params: path, password)
- 'unprotect': Unprotect workbook or sheet (required params: path, password)
- 'get': Get protection settings (required params: path)
- 'set_cell_locked': Set cell locked status (required params: path, range, locked)")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description(
            "Sheet index (0-based, optional). If not specified, the operation applies to the entire workbook structure.")]
        int? sheetIndex = null,
        [Description("Protection password (required for protect, optional for unprotect)")]
        string? password = null,
        [Description("Protect workbook structure (optional, for protect operation, default: false)")]
        bool protectWorkbook = false,
        [Description(
            "Protect workbook structure (optional, for protect operation when protectWorkbook is true, default: true)")]
        bool protectStructure = true,
        [Description(
            "Protect workbook windows (optional, for protect operation when protectWorkbook is true, default: false)")]
        bool protectWindows = false,
        [Description("Cell or range (e.g., 'A1' or 'A1:C5', required for set_cell_locked)")]
        string? range = null,
        [Description("Locked status (true = locked, false = unlocked, required for set_cell_locked)")]
        bool locked = false)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, sheetIndex, password, protectWorkbook, protectStructure,
            protectWindows, range, locked);

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
        int? sheetIndex,
        string? password,
        bool protectWorkbook,
        bool protectStructure,
        bool protectWindows,
        string? range,
        bool locked)
    {
        var parameters = new OperationParameters();
        if (sheetIndex.HasValue) parameters.Set("sheetIndex", sheetIndex.Value);

        return operation.ToLowerInvariant() switch
        {
            "protect" => BuildProtectParameters(parameters, password, protectWorkbook, protectStructure,
                protectWindows),
            "unprotect" => BuildUnprotectParameters(parameters, password),
            "get" => parameters,
            "set_cell_locked" => BuildSetCellLockedParameters(parameters, range, locked),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for the protect operation.
    /// </summary>
    /// <param name="parameters">Base parameters with optional sheet index.</param>
    /// <param name="password">The protection password.</param>
    /// <param name="protectWorkbook">Whether to protect workbook structure.</param>
    /// <param name="protectStructure">Whether to protect workbook structure when protectWorkbook is true.</param>
    /// <param name="protectWindows">Whether to protect workbook windows when protectWorkbook is true.</param>
    /// <returns>OperationParameters configured for protecting workbook or sheet.</returns>
    private static OperationParameters BuildProtectParameters(OperationParameters parameters, string? password,
        bool protectWorkbook, bool protectStructure, bool protectWindows)
    {
        if (password != null) parameters.Set("password", password);
        parameters.Set("protectWorkbook", protectWorkbook);
        parameters.Set("protectStructure", protectStructure);
        parameters.Set("protectWindows", protectWindows);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the unprotect operation.
    /// </summary>
    /// <param name="parameters">Base parameters with optional sheet index.</param>
    /// <param name="password">The protection password.</param>
    /// <returns>OperationParameters configured for unprotecting workbook or sheet.</returns>
    private static OperationParameters BuildUnprotectParameters(OperationParameters parameters, string? password)
    {
        if (password != null) parameters.Set("password", password);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the set cell locked operation.
    /// </summary>
    /// <param name="parameters">Base parameters with optional sheet index.</param>
    /// <param name="range">The cell or range to set locked status.</param>
    /// <param name="locked">The locked status.</param>
    /// <returns>OperationParameters configured for setting cell locked status.</returns>
    private static OperationParameters BuildSetCellLockedParameters(OperationParameters parameters, string? range,
        bool locked)
    {
        if (range != null) parameters.Set("range", range);
        parameters.Set("locked", locked);
        return parameters;
    }
}
