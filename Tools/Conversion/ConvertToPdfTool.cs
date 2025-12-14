using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Cells;
using Aspose.Slides;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

public class ConvertToPdfTool : IAsposeTool
{
    public string Description => @"Convert any document (Word, Excel, PowerPoint) to PDF.

Usage examples:
- Convert Word to PDF: convert_to_pdf(inputPath='doc.docx', outputPath='doc.pdf')
- Convert Excel to PDF: convert_to_pdf(inputPath='book.xlsx', outputPath='book.pdf')
- Convert PowerPoint to PDF: convert_to_pdf(inputPath='presentation.pptx', outputPath='presentation.pdf')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            inputPath = new
            {
                type = "string",
                description = "Input file path (required, supports Word, Excel, PowerPoint formats)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output PDF file path (required)"
            }
        },
        required = new[] { "inputPath", "outputPath" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var inputPath = ArgumentHelper.GetString(arguments, "inputPath", "inputPath");
        var outputPath = ArgumentHelper.GetString(arguments, "outputPath", "outputPath");

        SecurityHelper.ValidateFilePath(inputPath, "inputPath");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

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

