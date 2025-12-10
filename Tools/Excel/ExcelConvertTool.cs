using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelConvertTool : IAsposeTool
{
    public string Description => "Convert Excel to another format (PDF, HTML, CSV, etc.)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            inputPath = new
            {
                type = "string",
                description = "Input file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path"
            },
            format = new
            {
                type = "string",
                description = "Output format (pdf, html, csv, xlsx, xls, etc.)"
            }
        },
        required = new[] { "inputPath", "outputPath", "format" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var inputPath = arguments?["inputPath"]?.GetValue<string>() ?? throw new ArgumentException("inputPath is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? throw new ArgumentException("outputPath is required");
        var format = arguments?["format"]?.GetValue<string>()?.ToLower() ?? throw new ArgumentException("format is required");

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
            "xml" => SaveFormat.SpreadsheetML,
            _ => throw new ArgumentException($"Unsupported format: {format}")
        };

        workbook.Save(outputPath, saveFormat);

        return await Task.FromResult($"Excel converted from {inputPath} to {outputPath} ({format})");
    }
}

