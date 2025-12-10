using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

public class ExcelSplitWorkbookTool : IAsposeTool
{
    public string Description => "Split Excel workbook into multiple files (one sheet per file or by sheet selection)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            inputPath = new
            {
                type = "string",
                description = "Input workbook file path"
            },
            outputDirectory = new
            {
                type = "string",
                description = "Output directory path"
            },
            sheetIndices = new
            {
                type = "array",
                items = new { type = "number" },
                description = "Sheet indices to split (0-based, optional, if not provided splits all sheets)"
            },
            outputFileNamePattern = new
            {
                type = "string",
                description = "Output file name pattern, use {index} for sheet index, {name} for sheet name (optional, default: 'sheet_{name}.xlsx')"
            }
        },
        required = new[] { "inputPath", "outputDirectory" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var inputPath = arguments?["inputPath"]?.GetValue<string>() ?? throw new ArgumentException("inputPath is required");
        var outputDirectory = arguments?["outputDirectory"]?.GetValue<string>() ?? throw new ArgumentException("outputDirectory is required");
        var sheetIndicesArray = arguments?["sheetIndices"]?.AsArray();
        var fileNamePattern = arguments?["outputFileNamePattern"]?.GetValue<string>() ?? "sheet_{name}.xlsx";

        if (!Directory.Exists(outputDirectory))
        {
            Directory.CreateDirectory(outputDirectory);
        }

        using var workbook = new Workbook(inputPath);
        var totalSheets = workbook.Worksheets.Count;

        List<int> sheetIndices;
        if (sheetIndicesArray != null && sheetIndicesArray.Count > 0)
        {
            sheetIndices = sheetIndicesArray.Select(s => s?.GetValue<int>()).Where(s => s.HasValue).Select(s => s!.Value).ToList();
        }
        else
        {
            sheetIndices = Enumerable.Range(0, totalSheets).ToList();
        }

        var fileCount = 0;
        foreach (var sheetIndex in sheetIndices)
        {
            if (sheetIndex < 0 || sheetIndex >= totalSheets)
            {
                continue;
            }

            var sheet = workbook.Worksheets[sheetIndex];
            using var newWorkbook = new Workbook();
            newWorkbook.Worksheets.RemoveAt(0);
            var newSheet = newWorkbook.Worksheets.Add(sheet.Name);
            newSheet.Copy(sheet);

            var outputFileName = fileNamePattern
                .Replace("{index}", sheetIndex.ToString())
                .Replace("{name}", SecurityHelper.SanitizeFileName(sheet.Name));
            outputFileName = SecurityHelper.SanitizeFileName(outputFileName);
            var outputPath = Path.Combine(outputDirectory, outputFileName);
            newWorkbook.Save(outputPath);
            fileCount++;
        }

        return await Task.FromResult($"Split workbook into {fileCount} file(s) in: {outputDirectory}");
    }
}

