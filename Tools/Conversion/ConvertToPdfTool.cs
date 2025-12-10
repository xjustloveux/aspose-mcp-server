using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Cells;
using Aspose.Slides;

namespace AsposeMcpServer.Tools;

public class ConvertToPdfTool : IAsposeTool
{
    public string Description => "Convert any document (Word, Excel, PowerPoint) to PDF";

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
                description = "Output PDF file path"
            }
        },
        required = new[] { "inputPath", "outputPath" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var inputPath = arguments?["inputPath"]?.GetValue<string>() ?? throw new ArgumentException("inputPath is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? throw new ArgumentException("outputPath is required");

        var extension = Path.GetExtension(inputPath).ToLower();

        switch (extension)
        {
            case ".doc":
            case ".docx":
            case ".rtf":
            case ".odt":
                var wordDoc = new Document(inputPath);
                wordDoc.Save(outputPath, Aspose.Words.SaveFormat.Pdf);
                break;

            case ".xls":
            case ".xlsx":
            case ".csv":
            case ".ods":
                using (var workbook = new Workbook(inputPath))
                {
                    workbook.Save(outputPath, Aspose.Cells.SaveFormat.Pdf);
                }
                break;

            case ".ppt":
            case ".pptx":
            case ".odp":
                using (var presentation = new Presentation(inputPath))
                {
                    presentation.Save(outputPath, Aspose.Slides.Export.SaveFormat.Pdf);
                }
                break;

            default:
                throw new ArgumentException($"Unsupported file format: {extension}");
        }

        return await Task.FromResult($"Document converted to PDF: {outputPath}");
    }
}

