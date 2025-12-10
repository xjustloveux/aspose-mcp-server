using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordConvertTool : IAsposeTool
{
    public string Description => "Convert Word document to another format (PDF, HTML, DOCX, TXT, etc.)";

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
                description = "Output format (pdf, html, docx, txt, rtf, odt, etc.)"
            }
        },
        required = new[] { "inputPath", "outputPath", "format" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var inputPath = arguments?["inputPath"]?.GetValue<string>() ?? throw new ArgumentException("inputPath is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? throw new ArgumentException("outputPath is required");
        var format = arguments?["format"]?.GetValue<string>()?.ToLower() ?? throw new ArgumentException("format is required");

        var doc = new Document(inputPath);

        var saveFormat = format switch
        {
            "pdf" => SaveFormat.Pdf,
            "html" => SaveFormat.Html,
            "docx" => SaveFormat.Docx,
            "doc" => SaveFormat.Doc,
            "txt" => SaveFormat.Text,
            "rtf" => SaveFormat.Rtf,
            "odt" => SaveFormat.Odt,
            "epub" => SaveFormat.Epub,
            "xps" => SaveFormat.Xps,
            _ => throw new ArgumentException($"Unsupported format: {format}")
        };

        doc.Save(outputPath, saveFormat);

        return await Task.FromResult($"Document converted from {inputPath} to {outputPath} ({format})");
    }
}

