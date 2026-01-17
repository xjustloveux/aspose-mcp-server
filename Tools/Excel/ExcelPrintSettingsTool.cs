using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel print settings (print area, titles, page setup, etc.).
/// </summary>
[McpServerToolType]
public class ExcelPrintSettingsTool
{
    /// <summary>
    ///     Handler registry for print settings operations.
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
    ///     Initializes a new instance of the <see cref="ExcelPrintSettingsTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelPrintSettingsTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry =
            HandlerRegistry<Workbook>.CreateFromNamespace("AsposeMcpServer.Handlers.Excel.PrintSettings");
    }

    /// <summary>
    ///     Executes an Excel print settings operation (set_print_area, set_print_titles, set_page_setup, set_all).
    /// </summary>
    /// <param name="operation">The operation to perform: set_print_area, set_print_titles, set_page_setup, set_all.</param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sheetIndex">Sheet index (0-based, default: 0).</param>
    /// <param name="range">Print area range (e.g., 'A1:D10').</param>
    /// <param name="clearPrintArea">Clear print area (optional, for set_print_area, default: false).</param>
    /// <param name="rows">Rows to repeat on each printed page (e.g., '1:1').</param>
    /// <param name="columns">Columns to repeat on each printed page (e.g., 'A:A').</param>
    /// <param name="clearTitles">Clear print titles (optional, for set_print_titles, default: false).</param>
    /// <param name="orientation">Page orientation (optional, default: Portrait).</param>
    /// <param name="paperSize">Paper size (e.g., 'A4', 'Letter').</param>
    /// <param name="leftMargin">Left margin in inches.</param>
    /// <param name="rightMargin">Right margin in inches.</param>
    /// <param name="topMargin">Top margin in inches.</param>
    /// <param name="bottomMargin">Bottom margin in inches.</param>
    /// <param name="header">Header text for center section.</param>
    /// <param name="footer">Footer text for center section.</param>
    /// <param name="fitToPage">Enable fit to page mode.</param>
    /// <param name="fitToPagesWide">Number of pages wide to fit content.</param>
    /// <param name="fitToPagesTall">Number of pages tall to fit content.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "excel_print_settings")]
    [Description(
        @"Manage Excel print settings. Supports 4 operations: set_print_area, set_print_titles, set_page_setup, set_all.

Usage examples:
- Set print area: excel_print_settings(operation='set_print_area', path='book.xlsx', range='A1:D10')
- Set multiple print areas: excel_print_settings(operation='set_print_area', path='book.xlsx', range='A1:D10,F1:H10')
- Set print titles: excel_print_settings(operation='set_print_titles', path='book.xlsx', rows='1:1', columns='A:A')
- Set page setup: excel_print_settings(operation='set_page_setup', path='book.xlsx', orientation='Landscape', paperSize='A4')
- Set margins: excel_print_settings(operation='set_page_setup', path='book.xlsx', leftMargin=0.5, topMargin=0.75)
- Set fit to page: excel_print_settings(operation='set_all', path='book.xlsx', fitToPage=true, fitToPagesWide=1, fitToPagesTall=0)
- Set all: excel_print_settings(operation='set_all', path='book.xlsx', range='A1:D10', orientation='Portrait')")]
    public string Execute(
        [Description(@"Operation to perform.
- 'set_print_area': Set print area (required params: path, range or clearPrintArea)
- 'set_print_titles': Set print titles (required params: path)
- 'set_page_setup': Set page setup (required params: path)
- 'set_all': Set all print settings (required params: path)")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description(
            "Print area range. Supports single range (e.g., 'A1:D10') or multiple ranges separated by comma (e.g., 'A1:D10,F1:H10')")]
        string? range = null,
        [Description("Clear print area (optional, for set_print_area, default: false)")]
        bool clearPrintArea = false,
        [Description("Rows to repeat on each printed page (e.g., '1:1' for first row, '1:2' for first two rows)")]
        string? rows = null,
        [Description(
            "Columns to repeat on each printed page (e.g., 'A:A' for first column, 'A:B' for first two columns)")]
        string? columns = null,
        [Description("Clear print titles (optional, for set_print_titles, default: false)")]
        bool clearTitles = false,
        [Description("Page orientation (optional, default: Portrait)")]
        string? orientation = null,
        [Description(
            "Paper size. Supported values: A3, A4, A5, B4, B5, Letter, Legal, Tabloid, Executive (optional, default: A4)")]
        string? paperSize = null,
        [Description("Left margin in inches (optional, default: 0.7)")]
        double? leftMargin = null,
        [Description("Right margin in inches (optional, default: 0.7)")]
        double? rightMargin = null,
        [Description("Top margin in inches (optional, default: 0.75)")]
        double? topMargin = null,
        [Description("Bottom margin in inches (optional, default: 0.75)")]
        double? bottomMargin = null,
        [Description("Header text for center section (optional)")]
        string? header = null,
        [Description("Footer text for center section (optional)")]
        string? footer = null,
        [Description("Enable fit to page mode. When true, fitToPagesWide and fitToPagesTall are used.")]
        bool? fitToPage = null,
        [Description(
            "Number of pages wide to fit content (optional, default: 1 when fitToPage is true, use 0 for automatic)")]
        int? fitToPagesWide = null,
        [Description(
            "Number of pages tall to fit content (optional, default: 1 when fitToPage is true, use 0 for automatic)")]
        int? fitToPagesTall = null)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, sheetIndex, range, clearPrintArea, rows, columns, clearTitles,
            orientation, paperSize, leftMargin, rightMargin, topMargin, bottomMargin, header, footer,
            fitToPage, fitToPagesWide, fitToPagesTall);

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

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return $"{result}\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters(
        string operation,
        int sheetIndex,
        string? range,
        bool clearPrintArea,
        string? rows,
        string? columns,
        bool clearTitles,
        string? orientation,
        string? paperSize,
        double? leftMargin,
        double? rightMargin,
        double? topMargin,
        double? bottomMargin,
        string? header,
        string? footer,
        bool? fitToPage,
        int? fitToPagesWide,
        int? fitToPagesTall)
    {
        var parameters = new OperationParameters();
        parameters.Set("sheetIndex", sheetIndex);

        return operation.ToLowerInvariant() switch
        {
            "set_print_area" => BuildSetPrintAreaParameters(parameters, range, clearPrintArea),
            "set_print_titles" => BuildSetPrintTitlesParameters(parameters, rows, columns, clearTitles),
            "set_page_setup" => BuildSetPageSetupParameters(parameters, orientation, paperSize, leftMargin, rightMargin,
                topMargin, bottomMargin, header, footer, fitToPage, fitToPagesWide, fitToPagesTall),
            "set_all" => BuildSetAllParameters(parameters, range, rows, columns, orientation, paperSize, leftMargin,
                rightMargin, topMargin, bottomMargin, header, footer, fitToPage, fitToPagesWide, fitToPagesTall),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for the set print area operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="range">The print area range.</param>
    /// <param name="clearPrintArea">Whether to clear the print area.</param>
    /// <returns>OperationParameters configured for setting print area.</returns>
    private static OperationParameters BuildSetPrintAreaParameters(OperationParameters parameters, string? range,
        bool clearPrintArea)
    {
        if (range != null) parameters.Set("range", range);
        parameters.Set("clearPrintArea", clearPrintArea);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the set print titles operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="rows">The rows to repeat on each printed page.</param>
    /// <param name="columns">The columns to repeat on each printed page.</param>
    /// <param name="clearTitles">Whether to clear the print titles.</param>
    /// <returns>OperationParameters configured for setting print titles.</returns>
    private static OperationParameters BuildSetPrintTitlesParameters(OperationParameters parameters, string? rows,
        string? columns, bool clearTitles)
    {
        if (rows != null) parameters.Set("rows", rows);
        if (columns != null) parameters.Set("columns", columns);
        parameters.Set("clearTitles", clearTitles);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the set page setup operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="orientation">The page orientation.</param>
    /// <param name="paperSize">The paper size.</param>
    /// <param name="leftMargin">The left margin in inches.</param>
    /// <param name="rightMargin">The right margin in inches.</param>
    /// <param name="topMargin">The top margin in inches.</param>
    /// <param name="bottomMargin">The bottom margin in inches.</param>
    /// <param name="header">The header text.</param>
    /// <param name="footer">The footer text.</param>
    /// <param name="fitToPage">Whether to enable fit to page mode.</param>
    /// <param name="fitToPagesWide">The number of pages wide to fit content.</param>
    /// <param name="fitToPagesTall">The number of pages tall to fit content.</param>
    /// <returns>OperationParameters configured for setting page setup.</returns>
    private static OperationParameters BuildSetPageSetupParameters(
        OperationParameters parameters, string? orientation,
        string? paperSize, double? leftMargin, double? rightMargin, double? topMargin, double? bottomMargin,
        string? header, string? footer, bool? fitToPage, int? fitToPagesWide, int? fitToPagesTall)
    {
        if (orientation != null) parameters.Set("orientation", orientation);
        if (paperSize != null) parameters.Set("paperSize", paperSize);
        if (leftMargin.HasValue) parameters.Set("leftMargin", leftMargin.Value);
        if (rightMargin.HasValue) parameters.Set("rightMargin", rightMargin.Value);
        if (topMargin.HasValue) parameters.Set("topMargin", topMargin.Value);
        if (bottomMargin.HasValue) parameters.Set("bottomMargin", bottomMargin.Value);
        if (header != null) parameters.Set("header", header);
        if (footer != null) parameters.Set("footer", footer);
        if (fitToPage.HasValue) parameters.Set("fitToPage", fitToPage.Value);
        if (fitToPagesWide.HasValue) parameters.Set("fitToPagesWide", fitToPagesWide.Value);
        if (fitToPagesTall.HasValue) parameters.Set("fitToPagesTall", fitToPagesTall.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the set all print settings operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="range">The print area range.</param>
    /// <param name="rows">The rows to repeat on each printed page.</param>
    /// <param name="columns">The columns to repeat on each printed page.</param>
    /// <param name="orientation">The page orientation.</param>
    /// <param name="paperSize">The paper size.</param>
    /// <param name="leftMargin">The left margin in inches.</param>
    /// <param name="rightMargin">The right margin in inches.</param>
    /// <param name="topMargin">The top margin in inches.</param>
    /// <param name="bottomMargin">The bottom margin in inches.</param>
    /// <param name="header">The header text.</param>
    /// <param name="footer">The footer text.</param>
    /// <param name="fitToPage">Whether to enable fit to page mode.</param>
    /// <param name="fitToPagesWide">The number of pages wide to fit content.</param>
    /// <param name="fitToPagesTall">The number of pages tall to fit content.</param>
    /// <returns>OperationParameters configured for setting all print settings.</returns>
    private static OperationParameters BuildSetAllParameters(
        OperationParameters parameters, string? range,
        string? rows, string? columns, string? orientation, string? paperSize, double? leftMargin, double? rightMargin,
        double? topMargin, double? bottomMargin, string? header, string? footer, bool? fitToPage, int? fitToPagesWide,
        int? fitToPagesTall)
    {
        if (range != null) parameters.Set("range", range);
        if (rows != null) parameters.Set("rows", rows);
        if (columns != null) parameters.Set("columns", columns);
        if (orientation != null) parameters.Set("orientation", orientation);
        if (paperSize != null) parameters.Set("paperSize", paperSize);
        if (leftMargin.HasValue) parameters.Set("leftMargin", leftMargin.Value);
        if (rightMargin.HasValue) parameters.Set("rightMargin", rightMargin.Value);
        if (topMargin.HasValue) parameters.Set("topMargin", topMargin.Value);
        if (bottomMargin.HasValue) parameters.Set("bottomMargin", bottomMargin.Value);
        if (header != null) parameters.Set("header", header);
        if (footer != null) parameters.Set("footer", footer);
        if (fitToPage.HasValue) parameters.Set("fitToPage", fitToPage.Value);
        if (fitToPagesWide.HasValue) parameters.Set("fitToPagesWide", fitToPagesWide.Value);
        if (fitToPagesTall.HasValue) parameters.Set("fitToPagesTall", fitToPagesTall.Value);
        return parameters;
    }
}
