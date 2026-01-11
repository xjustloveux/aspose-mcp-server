using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Handlers.Excel.FileOperations;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for Excel file operations (create, convert, merge workbooks, split workbook).
/// </summary>
[McpServerToolType]
public class ExcelFileOperationsTool
{
    /// <summary>
    ///     Handler registry for file operations.
    /// </summary>
    private readonly HandlerRegistry<Workbook> _handlerRegistry;

    /// <summary>
    ///     The session identity accessor for session isolation.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     The document session manager for managing in-memory document sessions.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelFileOperationsTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public ExcelFileOperationsTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = ExcelFileOperationsHandlerRegistry.Create();
    }

    /// <summary>
    ///     Executes an Excel file operation (create, convert, merge, or split).
    /// </summary>
    /// <param name="operation">The operation to perform: create, convert, merge, or split.</param>
    /// <param name="sessionId">Session ID to read workbook from session (for convert, split).</param>
    /// <param name="path">File path (output path for create/merge, input path for split).</param>
    /// <param name="outputPath">Output file path (required for convert, optional for create).</param>
    /// <param name="inputPath">Input file path (required for convert/split).</param>
    /// <param name="outputDirectory">Output directory path (required for split).</param>
    /// <param name="sheetName">Initial sheet name (optional, for create).</param>
    /// <param name="format">Output format: pdf, html, csv, xlsx, xls, ods, txt, tsv (required for convert).</param>
    /// <param name="inputPaths">Array of input workbook file paths (required for merge).</param>
    /// <param name="mergeSheets">When true, merges data from sheets with same name by appending rows.</param>
    /// <param name="sheetIndices">Sheet indices to split (0-based, optional).</param>
    /// <param name="outputFileNamePattern">Output file name pattern with {index} and {name} placeholders.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when the operation is unknown or required parameters are missing.</exception>
    [McpServerTool(Name = "excel_file_operations")]
    [Description(@"Excel file operations. Supports 4 operations: create, convert, merge, split.

Usage examples:
- Create workbook: excel_file_operations(operation='create', path='new.xlsx')
- Convert format: excel_file_operations(operation='convert', inputPath='book.xlsx', outputPath='book.pdf', format='pdf')
- Convert from session: excel_file_operations(operation='convert', sessionId='sess_xxx', outputPath='book.pdf', format='pdf')
- Merge workbooks: excel_file_operations(operation='merge', path='merged.xlsx', inputPaths=['book1.xlsx', 'book2.xlsx'])
- Split workbook: excel_file_operations(operation='split', inputPath='book.xlsx', outputDirectory='output/')
- Split from session: excel_file_operations(operation='split', sessionId='sess_xxx', outputDirectory='output/')")]
    public string Execute(
        [Description(@"Operation to perform.
- 'create': Create a new workbook (required params: path or outputPath)
- 'convert': Convert workbook format (required params: inputPath or sessionId, outputPath, format)
- 'merge': Merge workbooks (required params: path or outputPath, inputPaths)
- 'split': Split workbook (required params: inputPath, path, or sessionId, outputDirectory)")]
        string operation,
        [Description("Session ID to read workbook from session (for convert, split)")]
        string? sessionId = null,
        [Description("File path (output path for create/merge operations, input path for split operation)")]
        string? path = null,
        [Description("Output file path (required for convert, optional for create)")]
        string? outputPath = null,
        [Description("Input file path (required for convert/split)")]
        string? inputPath = null,
        [Description("Output directory path (required for split)")]
        string? outputDirectory = null,
        [Description("Initial sheet name (optional, for create)")]
        string? sheetName = null,
        [Description("Output format: 'pdf', 'html', 'csv', 'xlsx', 'xls', 'ods', 'txt', 'tsv' (required for convert)")]
        string? format = null,
        [Description("Array of input workbook file paths (required for merge)")]
        string[]? inputPaths = null,
        [Description(
            "When true, merges data from sheets with the same name by appending rows (optional, for merge, default: false)")]
        bool mergeSheets = false,
        [Description("Sheet indices to split (0-based, optional, for split)")]
        int[]? sheetIndices = null,
        [Description(
            "Output file name pattern, use {index} for sheet index, {name} for sheet name (optional, for split, default: 'sheet_{name}.xlsx')")]
        string outputFileNamePattern = "sheet_{name}.xlsx")
    {
        var parameters = BuildParameters(operation, sessionId, path, outputPath, inputPath, outputDirectory,
            sheetName, format, inputPaths, mergeSheets, sheetIndices, outputFileNamePattern);

        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<Workbook>
        {
            Document = null!,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = inputPath ?? path,
            OutputPath = outputPath ?? path
        };

        return handler.Execute(operationContext, parameters);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters.
    /// </summary>
    private static OperationParameters BuildParameters(
        string operation,
        string? sessionId,
        string? path,
        string? outputPath,
        string? inputPath,
        string? outputDirectory,
        string? sheetName,
        string? format,
        string[]? inputPaths,
        bool mergeSheets,
        int[]? sheetIndices,
        string outputFileNamePattern)
    {
        var parameters = new OperationParameters();

        switch (operation.ToLowerInvariant())
        {
            case "create":
                if (path != null) parameters.Set("path", path);
                if (outputPath != null) parameters.Set("outputPath", outputPath);
                if (sheetName != null) parameters.Set("sheetName", sheetName);
                break;

            case "convert":
                if (inputPath != null) parameters.Set("inputPath", inputPath);
                if (sessionId != null) parameters.Set("sessionId", sessionId);
                if (outputPath != null) parameters.Set("outputPath", outputPath);
                if (format != null) parameters.Set("format", format);
                break;

            case "merge":
                if (path != null) parameters.Set("path", path);
                if (outputPath != null) parameters.Set("outputPath", outputPath);
                if (inputPaths != null) parameters.Set("inputPaths", inputPaths);
                parameters.Set("mergeSheets", mergeSheets);
                break;

            case "split":
                if (inputPath != null) parameters.Set("inputPath", inputPath);
                if (path != null) parameters.Set("path", path);
                if (sessionId != null) parameters.Set("sessionId", sessionId);
                if (outputDirectory != null) parameters.Set("outputDirectory", outputDirectory);
                if (sheetIndices != null) parameters.Set("sheetIndices", sheetIndices);
                parameters.Set("outputFileNamePattern", outputFileNamePattern);
                break;
        }

        return parameters;
    }
}
