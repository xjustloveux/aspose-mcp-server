using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for Excel file operations (create, merge workbooks, split workbook).
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Excel.FileOperations")]
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
        _handlerRegistry =
            HandlerRegistry<Workbook>.CreateFromNamespace("AsposeMcpServer.Handlers.Excel.FileOperations");
    }

    /// <summary>
    ///     Executes an Excel file operation (create, merge, or split).
    /// </summary>
    /// <param name="operation">The operation to perform: create, merge, or split.</param>
    /// <param name="sessionId">Session ID to read workbook from session (for split).</param>
    /// <param name="path">File path (output path for create/merge, input path for split).</param>
    /// <param name="outputPath">Output file path (optional for create).</param>
    /// <param name="inputPath">Input file path (required for split).</param>
    /// <param name="outputDirectory">Output directory path (required for split).</param>
    /// <param name="sheetName">Initial sheet name (optional, for create).</param>
    /// <param name="inputPaths">Array of input workbook file paths (required for merge).</param>
    /// <param name="mergeSheets">When true, merges data from sheets with same name by appending rows.</param>
    /// <param name="sheetIndices">Sheet indices to split (0-based, optional).</param>
    /// <param name="outputFileNamePattern">Output file name pattern with {index} and {name} placeholders.</param>
    /// <param name="progress">Optional progress reporter for long-running operations.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when the operation is unknown or required parameters are missing.</exception>
    [McpServerTool(
        Name = "excel_file_operations",
        Title = "Excel File Operations",
        Destructive = false,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(@"Excel file operations. Supports 3 operations: create, merge, split.
For workbook format conversion, use convert_document tool instead.

Usage examples:
- Create workbook: excel_file_operations(operation='create', path='new.xlsx')
- Merge workbooks: excel_file_operations(operation='merge', path='merged.xlsx', inputPaths=['book1.xlsx', 'book2.xlsx'])
- Split workbook: excel_file_operations(operation='split', inputPath='book.xlsx', outputDirectory='output/')
- Split from session: excel_file_operations(operation='split', sessionId='sess_xxx', outputDirectory='output/')")]
    public object Execute(
        [Description(@"Operation to perform.
- 'create': Create a new workbook (required params: path or outputPath)
- 'merge': Merge workbooks (required params: path or outputPath, inputPaths)
- 'split': Split workbook (required params: inputPath, path, or sessionId, outputDirectory)")]
        string operation,
        [Description("Session ID to read workbook from session (for split)")]
        string? sessionId = null,
        [Description("File path (output path for create/merge operations, input path for split operation)")]
        string? path = null,
        [Description("Output file path (optional for create)")]
        string? outputPath = null,
        [Description("Input file path (required for split)")]
        string? inputPath = null,
        [Description("Output directory path (required for split)")]
        string? outputDirectory = null,
        [Description("Initial sheet name (optional, for create)")]
        string? sheetName = null,
        [Description("Array of input workbook file paths (required for merge)")]
        string[]? inputPaths = null,
        [Description(
            "When true, merges data from sheets with the same name by appending rows (optional, for merge, default: false)")]
        bool mergeSheets = false,
        [Description("Sheet indices to split (0-based, optional, for split)")]
        int[]? sheetIndices = null,
        [Description(
            "Output file name pattern, use {index} for sheet index, {name} for sheet name (optional, for split, default: 'sheet_{name}.xlsx')")]
        string outputFileNamePattern = "sheet_{name}.xlsx",
        IProgress<ProgressNotificationValue>? progress = null)
    {
        var parameters = BuildParameters(operation, sessionId, path, outputPath, inputPath, outputDirectory,
            sheetName, inputPaths, mergeSheets, sheetIndices, outputFileNamePattern);

        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<Workbook>
        {
            Document = null!,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = inputPath ?? path,
            OutputPath = outputPath ?? path,
            Progress = progress
        };

        var result = handler.Execute(operationContext, parameters);
        return ResultHelper.FinalizeResult((dynamic)result, outputPath ?? path, sessionId);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters(
        string operation,
        string? sessionId,
        string? path,
        string? outputPath,
        string? inputPath,
        string? outputDirectory,
        string? sheetName,
        string[]? inputPaths,
        bool mergeSheets,
        int[]? sheetIndices,
        string outputFileNamePattern)
    {
        var parameters = new OperationParameters();

        return operation.ToLowerInvariant() switch
        {
            "create" => BuildCreateParameters(parameters, path, outputPath, sheetName),
            "merge" => BuildMergeParameters(parameters, path, outputPath, inputPaths, mergeSheets),
            "split" => BuildSplitParameters(parameters, inputPath, path, sessionId, outputDirectory, sheetIndices,
                outputFileNamePattern),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for the create workbook operation.
    /// </summary>
    /// <param name="parameters">Base parameters.</param>
    /// <param name="path">The output file path.</param>
    /// <param name="outputPath">Alternative output file path.</param>
    /// <param name="sheetName">The initial sheet name.</param>
    /// <returns>OperationParameters configured for creating workbook.</returns>
    private static OperationParameters BuildCreateParameters(OperationParameters parameters, string? path,
        string? outputPath, string? sheetName)
    {
        if (path != null) parameters.Set("path", path);
        if (outputPath != null) parameters.Set("outputPath", outputPath);
        if (sheetName != null) parameters.Set("sheetName", sheetName);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the merge workbooks operation.
    /// </summary>
    /// <param name="parameters">Base parameters.</param>
    /// <param name="path">The output file path.</param>
    /// <param name="outputPath">Alternative output file path.</param>
    /// <param name="inputPaths">The input file paths to merge.</param>
    /// <param name="mergeSheets">Whether to merge sheets with the same name.</param>
    /// <returns>OperationParameters configured for merging workbooks.</returns>
    private static OperationParameters BuildMergeParameters(OperationParameters parameters, string? path,
        string? outputPath, string[]? inputPaths, bool mergeSheets)
    {
        if (path != null) parameters.Set("path", path);
        if (outputPath != null) parameters.Set("outputPath", outputPath);
        if (inputPaths != null) parameters.Set("inputPaths", inputPaths);
        parameters.Set("mergeSheets", mergeSheets);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the split workbook operation.
    /// </summary>
    /// <param name="parameters">Base parameters.</param>
    /// <param name="inputPath">The input file path.</param>
    /// <param name="path">Alternative input file path.</param>
    /// <param name="sessionId">The session ID for in-memory workbook.</param>
    /// <param name="outputDirectory">The output directory path.</param>
    /// <param name="sheetIndices">The sheet indices to split.</param>
    /// <param name="outputFileNamePattern">The output file name pattern.</param>
    /// <returns>OperationParameters configured for splitting workbook.</returns>
    private static OperationParameters BuildSplitParameters(OperationParameters parameters, string? inputPath,
        string? path, string? sessionId, string? outputDirectory, int[]? sheetIndices, string outputFileNamePattern)
    {
        if (inputPath != null) parameters.Set("inputPath", inputPath);
        if (path != null) parameters.Set("path", path);
        if (sessionId != null) parameters.Set("sessionId", sessionId);
        if (outputDirectory != null) parameters.Set("outputDirectory", outputDirectory);
        if (sheetIndices != null) parameters.Set("sheetIndices", sheetIndices);
        parameters.Set("outputFileNamePattern", outputFileNamePattern);
        return parameters;
    }
}
