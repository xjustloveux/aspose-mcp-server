using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for Excel file operations (create, convert, merge workbooks, split workbook)
/// Merges: ExcelCreateTool, ExcelConvertTool, ExcelMergeWorkbooksTool, ExcelSplitWorkbookTool
/// </summary>
public class ExcelFileOperationsTool : IAsposeTool
{
    public string Description => "Excel file operations: create, convert, merge workbooks, or split workbook";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation to perform: 'create', 'convert', 'merge', 'split'",
                @enum = new[] { "create", "convert", "merge", "split" }
            },
            path = new
            {
                type = "string",
                description = "File path (output path for create, input path for convert/split)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (required for convert, optional for create, defaults to input path)"
            },
            inputPath = new
            {
                type = "string",
                description = "Input file path (required for convert/split)"
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
                description = "Output file name pattern, use {index} for sheet index, {name} for sheet name (optional, for split, default: 'sheet_{name}.xlsx')"
            }
        },
        required = new[] { "operation" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");

        return operation.ToLower() switch
        {
            "create" => await CreateWorkbookAsync(arguments),
            "convert" => await ConvertWorkbookAsync(arguments),
            "merge" => await MergeWorkbooksAsync(arguments),
            "split" => await SplitWorkbookAsync(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> CreateWorkbookAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? arguments?["outputPath"]?.GetValue<string>() ?? throw new ArgumentException("path or outputPath is required for create operation");
        var sheetName = arguments?["sheetName"]?.GetValue<string>();

        using var workbook = new Workbook();
        
        if (!string.IsNullOrEmpty(sheetName))
        {
            workbook.Worksheets[0].Name = sheetName;
        }

        workbook.Save(path);
        return await Task.FromResult($"Excel workbook created successfully at: {path}");
    }

    private async Task<string> ConvertWorkbookAsync(JsonObject? arguments)
    {
        var inputPath = arguments?["inputPath"]?.GetValue<string>() ?? throw new ArgumentException("inputPath is required for convert operation");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? throw new ArgumentException("outputPath is required for convert operation");
        var format = arguments?["format"]?.GetValue<string>()?.ToLower() ?? throw new ArgumentException("format is required for convert operation");

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
        return await Task.FromResult($"Workbook converted from {inputPath} to {outputPath} (format: {format})");
    }

    private async Task<string> MergeWorkbooksAsync(JsonObject? arguments)
    {
        var outputPath = arguments?["path"]?.GetValue<string>() ?? arguments?["outputPath"]?.GetValue<string>() ?? throw new ArgumentException("path or outputPath is required for merge operation");
        var inputPathsArray = arguments?["inputPaths"]?.AsArray() ?? throw new ArgumentException("inputPaths is required for merge operation");
        var mergeSheets = arguments?["mergeSheets"]?.GetValue<bool?>() ?? false;

        // Validate array size
        SecurityHelper.ValidateArraySize(inputPathsArray, "inputPaths");

        if (inputPathsArray.Count == 0)
        {
            throw new ArgumentException("At least one input path is required");
        }

        var inputPaths = inputPathsArray.Select(p => p?.GetValue<string>()).Where(p => !string.IsNullOrEmpty(p)).ToList();
        if (inputPaths.Count == 0)
        {
            throw new ArgumentException("No valid input paths provided");
        }

        // Validate all input paths
        foreach (var inputPath in inputPaths)
        {
            SecurityHelper.ValidateFilePath(inputPath!, "inputPaths");
        }

        // Validate output path
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        using var targetWorkbook = new Workbook(inputPaths[0]);

        for (int i = 1; i < inputPaths.Count; i++)
        {
            using var sourceWorkbook = new Workbook(inputPaths[i]);
            
            foreach (Worksheet sourceSheet in sourceWorkbook.Worksheets)
            {
                if (mergeSheets)
                {
                    var existingSheet = targetWorkbook.Worksheets[sourceSheet.Name];
                    if (existingSheet != null)
                    {
                        var lastRow = existingSheet.Cells.MaxDataRow + 1;
                        var sourceRange = sourceSheet.Cells.CreateRange(0, 0, sourceSheet.Cells.MaxDataRow + 1, sourceSheet.Cells.MaxDataColumn + 1);
                        var destRange = existingSheet.Cells.CreateRange(lastRow, 0, sourceSheet.Cells.MaxDataRow + 1, sourceSheet.Cells.MaxDataColumn + 1);
                        destRange.Copy(sourceRange, new PasteOptions { PasteType = PasteType.All });
                    }
                    else
                    {
                        targetWorkbook.Worksheets.Add(sourceSheet.Name);
                        var newSheet = targetWorkbook.Worksheets[targetWorkbook.Worksheets.Count - 1];
                        var sourceRange = sourceSheet.Cells.CreateRange(0, 0, sourceSheet.Cells.MaxDataRow + 1, sourceSheet.Cells.MaxDataColumn + 1);
                        var destRange = newSheet.Cells.CreateRange(0, 0, sourceSheet.Cells.MaxDataRow + 1, sourceSheet.Cells.MaxDataColumn + 1);
                        destRange.Copy(sourceRange, new PasteOptions { PasteType = PasteType.All });
                    }
                }
                else
                {
                    targetWorkbook.Worksheets.Add(sourceSheet.Name);
                    var newSheet = targetWorkbook.Worksheets[targetWorkbook.Worksheets.Count - 1];
                    var sourceRange = sourceSheet.Cells.CreateRange(0, 0, sourceSheet.Cells.MaxDataRow + 1, sourceSheet.Cells.MaxDataColumn + 1);
                    var destRange = newSheet.Cells.CreateRange(0, 0, sourceSheet.Cells.MaxDataRow + 1, sourceSheet.Cells.MaxDataColumn + 1);
                    destRange.Copy(sourceRange, new PasteOptions { PasteType = PasteType.All });
                }
            }
        }

        targetWorkbook.Save(outputPath);
        return await Task.FromResult($"Merged {inputPaths.Count} workbooks into {outputPath}");
    }

    private async Task<string> SplitWorkbookAsync(JsonObject? arguments)
    {
        var inputPath = arguments?["inputPath"]?.GetValue<string>() ?? arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("inputPath or path is required for split operation");
        var outputDirectory = arguments?["outputDirectory"]?.GetValue<string>() ?? throw new ArgumentException("outputDirectory is required for split operation");
        var sheetIndicesArray = arguments?["sheetIndices"]?.AsArray();
        var fileNamePattern = arguments?["outputFileNamePattern"]?.GetValue<string>() ?? "sheet_{name}.xlsx";

        if (!Directory.Exists(outputDirectory))
        {
            Directory.CreateDirectory(outputDirectory);
        }

        using var sourceWorkbook = new Workbook(inputPath);
        var sheetIndices = sheetIndicesArray != null 
            ? sheetIndicesArray.Select(s => s?.GetValue<int>()).Where(s => s.HasValue).Select(s => s!.Value).ToList()
            : Enumerable.Range(0, sourceWorkbook.Worksheets.Count).ToList();

        var splitFiles = new List<string>();

        foreach (var sheetIndex in sheetIndices)
        {
            if (sheetIndex < 0 || sheetIndex >= sourceWorkbook.Worksheets.Count)
            {
                continue;
            }

            var worksheet = sourceWorkbook.Worksheets[sheetIndex];
            var fileName = fileNamePattern
                .Replace("{index}", sheetIndex.ToString())
                .Replace("{name}", worksheet.Name);
            var outputPath = Path.Combine(outputDirectory, fileName);

            using var newWorkbook = new Workbook();
            newWorkbook.Worksheets.RemoveAt(0);
            var newSheet = newWorkbook.Worksheets.Add(worksheet.Name);
            var sourceRange = worksheet.Cells.CreateRange(0, 0, worksheet.Cells.MaxDataRow + 1, worksheet.Cells.MaxDataColumn + 1);
            var destRange = newSheet.Cells.CreateRange(0, 0, worksheet.Cells.MaxDataRow + 1, worksheet.Cells.MaxDataColumn + 1);
            destRange.Copy(sourceRange, new PasteOptions { PasteType = PasteType.All });
            newWorkbook.Save(outputPath);
            splitFiles.Add(outputPath);
        }

        return await Task.FromResult($"Split workbook into {splitFiles.Count} files:\n{string.Join("\n", splitFiles)}");
    }
}

