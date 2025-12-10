using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelMergeWorkbooksTool : IAsposeTool
{
    public string Description => "Merge multiple Excel workbooks into one";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            outputPath = new
            {
                type = "string",
                description = "Output workbook file path"
            },
            inputPaths = new
            {
                type = "array",
                items = new { type = "string" },
                description = "Array of input workbook file paths"
            },
            mergeSheets = new
            {
                type = "boolean",
                description = "Merge sheets with same names (optional, default: false)"
            }
        },
        required = new[] { "outputPath", "inputPaths" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? throw new ArgumentException("outputPath is required");
        var inputPathsArray = arguments?["inputPaths"]?.AsArray() ?? throw new ArgumentException("inputPaths is required");
        var mergeSheets = arguments?["mergeSheets"]?.GetValue<bool?>() ?? false;

        if (inputPathsArray.Count == 0)
        {
            throw new ArgumentException("At least one input path is required");
        }

        var inputPaths = inputPathsArray.Select(p => p?.GetValue<string>()).Where(p => !string.IsNullOrEmpty(p)).ToList();
        if (inputPaths.Count == 0)
        {
            throw new ArgumentException("No valid input paths provided");
        }

        var masterWorkbook = new Workbook(inputPaths[0]!);

        for (int i = 1; i < inputPaths.Count; i++)
        {
            var inputPath = inputPaths[i];
            if (string.IsNullOrEmpty(inputPath) || !File.Exists(inputPath))
            {
                continue;
            }

            var sourceWorkbook = new Workbook(inputPath);
            
            foreach (Worksheet sourceSheet in sourceWorkbook.Worksheets)
            {
                if (mergeSheets)
                {
                    // Try to find sheet with same name
                    var existingSheet = masterWorkbook.Worksheets[sourceSheet.Name];
                    if (existingSheet != null)
                    {
                        // Append data to existing sheet
                        var maxRow = existingSheet.Cells.MaxDataRow + 1;
                        var sourceMaxRow = sourceSheet.Cells.MaxDataRow;
                        var sourceMaxCol = sourceSheet.Cells.MaxDataColumn;
                        if (sourceMaxRow >= 0 && sourceMaxCol >= 0)
                        {
                            for (int row = 0; row <= sourceMaxRow; row++)
                            {
                                for (int col = 0; col <= sourceMaxCol; col++)
                                {
                                    var sourceCell = sourceSheet.Cells[row, col];
                                    var targetCell = existingSheet.Cells[maxRow + row, col];
                                    targetCell.PutValue(sourceCell.Value);
                                    try
                                    {
                                        targetCell.SetStyle(sourceCell.GetStyle());
                                    }
                                    catch
                                    {
                                        // Ignore style copy errors
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        // Add as new sheet
                        var newSheet = masterWorkbook.Worksheets.Add(sourceSheet.Name);
                        newSheet.Copy(sourceSheet);
                    }
                }
                else
                {
                    // Add as new sheet (rename if duplicate)
                    var sheetName = sourceSheet.Name;
                    var counter = 1;
                    while (masterWorkbook.Worksheets[sheetName] != null)
                    {
                        sheetName = $"{sourceSheet.Name}_{counter}";
                        counter++;
                    }
                    var newSheet = masterWorkbook.Worksheets.Add(sheetName);
                    newSheet.Copy(sourceSheet);
                }
            }
        }

        masterWorkbook.Save(outputPath);
        return await Task.FromResult($"Merged {inputPaths.Count} workbooks into: {outputPath} (Total sheets: {masterWorkbook.Worksheets.Count})");
    }
}

