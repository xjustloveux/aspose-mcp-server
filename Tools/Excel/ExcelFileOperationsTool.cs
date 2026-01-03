using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for Excel file operations (create, convert, merge workbooks, split workbook).
///     Merges: ExcelCreateTool, ExcelConvertTool, ExcelMergeWorkbooksTool, ExcelSplitWorkbookTool.
/// </summary>
[McpServerToolType]
public class ExcelFileOperationsTool
{
    [McpServerTool(Name = "excel_file_operations")]
    [Description(@"Excel file operations. Supports 4 operations: create, convert, merge, split.

Usage examples:
- Create workbook: excel_file_operations(operation='create', path='new.xlsx')
- Convert format: excel_file_operations(operation='convert', inputPath='book.xlsx', outputPath='book.pdf', format='pdf')
- Merge workbooks: excel_file_operations(operation='merge', path='merged.xlsx', inputPaths=['book1.xlsx', 'book2.xlsx'])
- Split workbook: excel_file_operations(operation='split', inputPath='book.xlsx', outputDirectory='output/')")]
    public string Execute(
        [Description(@"Operation to perform.
- 'create': Create a new workbook (required params: path or outputPath)
- 'convert': Convert workbook format (required params: inputPath, outputPath, format)
- 'merge': Merge workbooks (required params: path or outputPath, inputPaths)
- 'split': Split workbook (required params: inputPath or path, outputDirectory)")]
        string operation,
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
        return operation.ToLower() switch
        {
            "create" => CreateWorkbook(path, outputPath, sheetName),
            "convert" => ConvertWorkbook(inputPath, outputPath, format),
            "merge" => MergeWorkbooks(path, outputPath, inputPaths, mergeSheets),
            "split" => SplitWorkbook(inputPath, path, outputDirectory, sheetIndices, outputFileNamePattern),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Creates a new Excel workbook.
    /// </summary>
    /// <param name="path">The file path for the new workbook.</param>
    /// <param name="outputPath">Alternative output file path.</param>
    /// <param name="sheetName">Optional name for the initial worksheet.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when neither path nor outputPath is provided.</exception>
    private static string CreateWorkbook(string? path, string? outputPath, string? sheetName)
    {
        var targetPath = path ?? outputPath ?? throw new ArgumentException("path or outputPath is required");
        SecurityHelper.ValidateFilePath(targetPath, allowAbsolutePaths: true);

        using var workbook = new Workbook();

        if (!string.IsNullOrEmpty(sheetName))
            workbook.Worksheets[0].Name = sheetName;

        workbook.Save(targetPath);
        return $"Excel workbook created successfully. Output: {targetPath}";
    }

    /// <summary>
    ///     Converts an Excel workbook to a different format.
    /// </summary>
    /// <param name="inputPath">The source workbook file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="format">The target format (pdf, html, csv, xlsx, xls, ods, txt, tsv).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when inputPath, outputPath, or format is not provided, or when the format is
    ///     unsupported.
    /// </exception>
    private static string ConvertWorkbook(string? inputPath, string? outputPath, string? format)
    {
        if (string.IsNullOrEmpty(inputPath))
            throw new ArgumentException("inputPath is required for convert operation");
        if (string.IsNullOrEmpty(outputPath))
            throw new ArgumentException("outputPath is required for convert operation");
        if (string.IsNullOrEmpty(format))
            throw new ArgumentException("format is required for convert operation");

        using var workbook = new Workbook(inputPath);

        var saveFormat = format.ToLower() switch
        {
            "pdf" => SaveFormat.Pdf,
            "html" => SaveFormat.Html,
            "csv" => SaveFormat.Csv,
            "xlsx" => SaveFormat.Xlsx,
            "xls" => SaveFormat.Excel97To2003,
            "ods" => SaveFormat.Ods,
            "txt" => SaveFormat.TabDelimited,
            "tsv" => SaveFormat.TabDelimited,
            _ => throw new ArgumentException($"Unsupported format: {format}")
        };

        workbook.Save(outputPath, saveFormat);
        return $"Workbook converted to {format} format. Output: {outputPath}";
    }

    /// <summary>
    ///     Merges multiple Excel workbooks into one.
    /// </summary>
    /// <param name="path">The output file path.</param>
    /// <param name="outputPath">Alternative output file path.</param>
    /// <param name="inputPaths">Array of input workbook file paths to merge.</param>
    /// <param name="mergeSheets">Whether to merge data from sheets with the same name by appending rows.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when neither path nor outputPath is provided, or when no valid input paths
    ///     are provided.
    /// </exception>
    private static string MergeWorkbooks(string? path, string? outputPath, string[]? inputPaths, bool mergeSheets)
    {
        var targetPath = path ?? outputPath ?? throw new ArgumentException("path or outputPath is required");

        if (inputPaths == null || inputPaths.Length == 0)
            throw new ArgumentException("At least one input path is required");

        var validPaths = inputPaths.Where(p => !string.IsNullOrEmpty(p)).ToList();
        if (validPaths.Count == 0)
            throw new ArgumentException("No valid input paths provided");

        foreach (var inputPath in validPaths)
            SecurityHelper.ValidateFilePath(inputPath, "inputPaths", true);

        SecurityHelper.ValidateFilePath(targetPath, "outputPath", true);

        using var targetWorkbook = new Workbook(validPaths[0]);

        for (var i = 1; i < validPaths.Count; i++)
        {
            using var sourceWorkbook = new Workbook(validPaths[i]);

            if (mergeSheets)
                foreach (var sourceSheet in sourceWorkbook.Worksheets)
                {
                    var existingSheet = targetWorkbook.Worksheets[sourceSheet.Name];
                    if (existingSheet != null)
                    {
                        var lastRow = existingSheet.Cells.MaxDataRow + 1;
                        var sourceMaxRow = sourceSheet.Cells.MaxDataRow;
                        var sourceMaxCol = sourceSheet.Cells.MaxDataColumn;

                        if (sourceMaxRow >= 0 && sourceMaxCol >= 0)
                        {
                            var sourceRange =
                                sourceSheet.Cells.CreateRange(0, 0, sourceMaxRow + 1, sourceMaxCol + 1);
                            var destRange =
                                existingSheet.Cells.CreateRange(lastRow, 0, sourceMaxRow + 1, sourceMaxCol + 1);
                            destRange.Copy(sourceRange, new PasteOptions { PasteType = PasteType.All });
                        }
                    }
                    else
                    {
                        targetWorkbook.Worksheets.Add(sourceSheet.Name);
                        var newSheet = targetWorkbook.Worksheets[^1];
                        var sourceMaxRow = sourceSheet.Cells.MaxDataRow;
                        var sourceMaxCol = sourceSheet.Cells.MaxDataColumn;

                        if (sourceMaxRow >= 0 && sourceMaxCol >= 0)
                        {
                            var sourceRange =
                                sourceSheet.Cells.CreateRange(0, 0, sourceMaxRow + 1, sourceMaxCol + 1);
                            var destRange = newSheet.Cells.CreateRange(0, 0, sourceMaxRow + 1, sourceMaxCol + 1);
                            destRange.Copy(sourceRange, new PasteOptions { PasteType = PasteType.All });
                        }
                    }
                }
            else
                targetWorkbook.Combine(sourceWorkbook);
        }

        targetWorkbook.Save(targetPath);
        return $"Merged {validPaths.Count} workbooks successfully. Output: {targetPath}";
    }

    /// <summary>
    ///     Splits an Excel workbook into separate files, one per worksheet.
    /// </summary>
    /// <param name="inputPath">The source workbook file path.</param>
    /// <param name="path">Alternative source file path.</param>
    /// <param name="outputDirectory">The directory to save split files.</param>
    /// <param name="sheetIndices">Optional array of specific sheet indices to split.</param>
    /// <param name="fileNamePattern">The output file name pattern with {index} and {name} placeholders.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when neither inputPath nor path is provided, or when outputDirectory is not
    ///     provided.
    /// </exception>
    private static string SplitWorkbook(string? inputPath, string? path, string? outputDirectory, int[]? sheetIndices,
        string fileNamePattern)
    {
        var sourcePath = inputPath ?? path ?? throw new ArgumentException("inputPath or path is required");
        if (string.IsNullOrEmpty(outputDirectory))
            throw new ArgumentException("outputDirectory is required for split operation");

        if (!Directory.Exists(outputDirectory))
            Directory.CreateDirectory(outputDirectory);

        using var sourceWorkbook = new Workbook(sourcePath);

        var indicesToSplit = sheetIndices is { Length: > 0 }
            ? sheetIndices.Distinct().ToList()
            : Enumerable.Range(0, sourceWorkbook.Worksheets.Count).ToList();

        List<string> splitFiles = [];

        foreach (var sheetIndex in indicesToSplit)
        {
            if (sheetIndex < 0 || sheetIndex >= sourceWorkbook.Worksheets.Count)
                continue;

            var worksheet = sourceWorkbook.Worksheets[sheetIndex];
            var fileName = fileNamePattern
                .Replace("{index}", sheetIndex.ToString())
                .Replace("{name}", worksheet.Name);
            var outputFilePath = Path.Combine(outputDirectory, fileName);

            using var newWorkbook = new Workbook();
            newWorkbook.Worksheets.RemoveAt(0);
            newWorkbook.Worksheets.Add(worksheet.Name);
            var newSheet = newWorkbook.Worksheets[0];
            var maxRow = worksheet.Cells.MaxDataRow;
            var maxCol = worksheet.Cells.MaxDataColumn;

            if (maxRow >= 0 && maxCol >= 0)
            {
                var sourceRange = worksheet.Cells.CreateRange(0, 0, maxRow + 1, maxCol + 1);
                var destRange = newSheet.Cells.CreateRange(0, 0, maxRow + 1, maxCol + 1);
                destRange.Copy(sourceRange, new PasteOptions { PasteType = PasteType.All });
            }

            newWorkbook.Save(outputFilePath);
            splitFiles.Add(outputFilePath);
        }

        return $"Split workbook into {splitFiles.Count} files. Output: {outputDirectory}";
    }
}