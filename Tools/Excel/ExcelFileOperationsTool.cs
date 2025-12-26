using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for Excel file operations (create, convert, merge workbooks, split workbook)
///     Merges: ExcelCreateTool, ExcelConvertTool, ExcelMergeWorkbooksTool, ExcelSplitWorkbookTool
/// </summary>
public class ExcelFileOperationsTool : IAsposeTool
{
    public string Description => @"Excel file operations. Supports 4 operations: create, convert, merge, split.

Usage examples:
- Create workbook: excel_file_operations(operation='create', path='new.xlsx')
- Convert format: excel_file_operations(operation='convert', inputPath='book.xlsx', outputPath='book.pdf', format='pdf')
- Merge workbooks: excel_file_operations(operation='merge', path='merged.xlsx', inputPaths=['book1.xlsx', 'book2.xlsx'])
- Split workbook: excel_file_operations(operation='split', inputPath='book.xlsx', outputDirectory='output/')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'create': Create a new workbook (required params: path or outputPath)
- 'convert': Convert workbook format (required params: inputPath, outputPath, format)
- 'merge': Merge workbooks (required params: path or outputPath, inputPaths)
- 'split': Split workbook (required params: inputPath or path, outputDirectory)",
                @enum = new[] { "create", "convert", "merge", "split" }
            },
            path = new
            {
                type = "string",
                description =
                    "File path (output path for create/merge operations, input path for split operation, can be used instead of inputPath/outputPath)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (required for convert, optional for create, defaults to input path)"
            },
            inputPath = new
            {
                type = "string",
                description = "Input file path (required for convert/split, can also use 'path' for split)"
            },
            outputDirectory = new
            {
                type = "string",
                description = "Output directory path (required for split)"
            },
            sheetName = new
            {
                type = "string",
                description = "Initial sheet name (optional, for create)"
            },
            format = new
            {
                type = "string",
                description = "Output format (pdf, html, csv, xlsx, xls, etc., required for convert)"
            },
            inputPaths = new
            {
                type = "array",
                items = new { type = "string" },
                description = "Array of input workbook file paths (required for merge)"
            },
            mergeSheets = new
            {
                type = "boolean",
                description = "Merge sheets with same names (optional, for merge, default: false)"
            },
            sheetIndices = new
            {
                type = "array",
                items = new { type = "number" },
                description = "Sheet indices to split (0-based, optional, for split, if not provided splits all sheets)"
            },
            outputFileNamePattern = new
            {
                type = "string",
                description =
                    "Output file name pattern, use {index} for sheet index, {name} for sheet name (optional, for split, default: 'sheet_{name}.xlsx')"
            }
        },
        required = new[] { "operation" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");

        return operation.ToLower() switch
        {
            "create" => await CreateWorkbookAsync(arguments),
            "convert" => await ConvertWorkbookAsync(arguments),
            "merge" => await MergeWorkbooksAsync(arguments),
            "split" => await SplitWorkbookAsync(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Creates a new workbook
    /// </summary>
    /// <param name="arguments">JSON arguments containing path or outputPath, optional sheetName</param>
    /// <returns>Success message with file path</returns>
    private Task<string> CreateWorkbookAsync(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var path = ArgumentHelper.GetString(arguments, "path", "outputPath", "path or outputPath");
            var sheetName = ArgumentHelper.GetStringNullable(arguments, "sheetName");

            using var workbook = new Workbook();

            if (!string.IsNullOrEmpty(sheetName)) workbook.Worksheets[0].Name = sheetName;

            // For create operation, path is the output path
            SecurityHelper.ValidateFilePath(path, allowAbsolutePaths: true);
            workbook.Save(path);
            return $"Excel workbook created successfully. Output: {path}";
        });
    }

    /// <summary>
    ///     Converts workbook to another format
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, outputPath, format</param>
    /// <returns>Success message with output path</returns>
    private Task<string> ConvertWorkbookAsync(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var inputPath = ArgumentHelper.GetString(arguments, "inputPath");
            var outputPath = ArgumentHelper.GetString(arguments, "outputPath");
            var format = ArgumentHelper.GetString(arguments, "format").ToLower();

            using var workbook = new Workbook(inputPath);

            var saveFormat = format switch
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
        });
    }

    /// <summary>
    ///     Merges multiple workbooks into one
    /// </summary>
    /// <param name="arguments">JSON arguments containing sourcePaths array and outputPath</param>
    /// <returns>Success message with merged file path</returns>
    private Task<string> MergeWorkbooksAsync(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var outputPath = ArgumentHelper.GetString(arguments, "path", "outputPath", "path or outputPath");
            var inputPathsArray = ArgumentHelper.GetArray(arguments, "inputPaths");
            var mergeSheets = ArgumentHelper.GetBool(arguments, "mergeSheets", false);

            // Validate array size
            SecurityHelper.ValidateArraySize(inputPathsArray, "inputPaths");

            if (inputPathsArray.Count == 0) throw new ArgumentException("At least one input path is required");

            var inputPaths = inputPathsArray.Select(p => p?.GetValue<string>()).Where(p => !string.IsNullOrEmpty(p))
                .ToList();
            if (inputPaths.Count == 0) throw new ArgumentException("No valid input paths provided");

            // Validate all input paths
            foreach (var inputPath in inputPaths) SecurityHelper.ValidateFilePath(inputPath!, "inputPaths", true);

            // Validate output path
            SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

            using var targetWorkbook = new Workbook(inputPaths[0]);

            for (var i = 1; i < inputPaths.Count; i++)
            {
                using var sourceWorkbook = new Workbook(inputPaths[i]);

                foreach (var sourceSheet in sourceWorkbook.Worksheets)
                    if (mergeSheets)
                    {
                        var existingSheet = targetWorkbook.Worksheets[sourceSheet.Name];
                        if (existingSheet != null)
                        {
                            var lastRow = existingSheet.Cells.MaxDataRow + 1;
                            var sourceRange = sourceSheet.Cells.CreateRange(0, 0, sourceSheet.Cells.MaxDataRow + 1,
                                sourceSheet.Cells.MaxDataColumn + 1);
                            var destRange = existingSheet.Cells.CreateRange(lastRow, 0,
                                sourceSheet.Cells.MaxDataRow + 1,
                                sourceSheet.Cells.MaxDataColumn + 1);
                            destRange.Copy(sourceRange, new PasteOptions { PasteType = PasteType.All });
                        }
                        else
                        {
                            targetWorkbook.Worksheets.Add(sourceSheet.Name);
                            var newSheet = targetWorkbook.Worksheets[^1];
                            var sourceRange = sourceSheet.Cells.CreateRange(0, 0, sourceSheet.Cells.MaxDataRow + 1,
                                sourceSheet.Cells.MaxDataColumn + 1);
                            var destRange = newSheet.Cells.CreateRange(0, 0, sourceSheet.Cells.MaxDataRow + 1,
                                sourceSheet.Cells.MaxDataColumn + 1);
                            destRange.Copy(sourceRange, new PasteOptions { PasteType = PasteType.All });
                        }
                    }
                    else
                    {
                        // Check if sheet name already exists and rename if necessary
                        var sheetName = sourceSheet.Name;
                        var existingNames = targetWorkbook.Worksheets
                            .Select(ws => ws.Name).ToList();
                        var nameCounter = 1;
                        while (existingNames.Contains(sheetName))
                        {
                            sheetName = $"{sourceSheet.Name}_{nameCounter}";
                            nameCounter++;
                        }

                        targetWorkbook.Worksheets.Add(sheetName);
                        var newSheet = targetWorkbook.Worksheets[^1];
                        var maxRow = sourceSheet.Cells.MaxDataRow;
                        var maxCol = sourceSheet.Cells.MaxDataColumn;

                        // Only copy if there's data
                        if (maxRow >= 0 && maxCol >= 0)
                        {
                            var sourceRange = sourceSheet.Cells.CreateRange(0, 0, maxRow + 1, maxCol + 1);
                            var destRange = newSheet.Cells.CreateRange(0, 0, maxRow + 1, maxCol + 1);
                            destRange.Copy(sourceRange, new PasteOptions { PasteType = PasteType.All });
                        }
                    }
            }

            targetWorkbook.Save(outputPath);
            return $"Merged {inputPaths.Count} workbooks successfully. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Splits workbook into multiple files (one file per sheet)
    /// </summary>
    /// <param name="arguments">
    ///     JSON arguments containing inputPath, outputDirectory, optional sheetIndices and
    ///     outputFileNamePattern
    /// </param>
    /// <returns>Success message with split file count</returns>
    private Task<string> SplitWorkbookAsync(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var inputPath = ArgumentHelper.GetString(arguments, "inputPath", "path", "inputPath or path");
            var outputDirectory = ArgumentHelper.GetString(arguments, "outputDirectory");
            var sheetIndicesArray = ArgumentHelper.GetArray(arguments, "sheetIndices", false);
            var fileNamePattern = ArgumentHelper.GetString(arguments, "outputFileNamePattern", "sheet_{name}.xlsx");

            if (!Directory.Exists(outputDirectory)) Directory.CreateDirectory(outputDirectory);

            using var sourceWorkbook = new Workbook(inputPath);
            var sheetIndices = sheetIndicesArray != null
                ? sheetIndicesArray.Select(s => s?.GetValue<int>()).Where(s => s.HasValue).Select(s => s!.Value)
                    .ToList()
                : Enumerable.Range(0, sourceWorkbook.Worksheets.Count).ToList();

            var splitFiles = new List<string>();

            foreach (var sheetIndex in sheetIndices)
            {
                if (sheetIndex < 0 || sheetIndex >= sourceWorkbook.Worksheets.Count) continue;

                var worksheet = sourceWorkbook.Worksheets[sheetIndex];
                var fileName = fileNamePattern
                    .Replace("{index}", sheetIndex.ToString())
                    .Replace("{name}", worksheet.Name);
                var outputPath = Path.Combine(outputDirectory, fileName);

                using var newWorkbook = new Workbook();
                newWorkbook.Worksheets.RemoveAt(0);
                var newSheet = newWorkbook.Worksheets.Add(worksheet.Name);
                var maxRow = worksheet.Cells.MaxDataRow;
                var maxCol = worksheet.Cells.MaxDataColumn;

                // Only copy if there's data (maxRow and maxCol are >= 0)
                if (maxRow >= 0 && maxCol >= 0)
                {
                    var sourceRange = worksheet.Cells.CreateRange(0, 0, maxRow + 1, maxCol + 1);
                    var destRange = newSheet.Cells.CreateRange(0, 0, maxRow + 1, maxCol + 1);
                    destRange.Copy(sourceRange, new PasteOptions { PasteType = PasteType.All });
                }

                newWorkbook.Save(outputPath);
                splitFiles.Add(outputPath);
            }

            return $"Split workbook into {splitFiles.Count} files. Output: {outputDirectory}";
        });
    }
}