using System.Text.Json.Nodes;
using Aspose.Cells;
using Aspose.Slides;
using Aspose.Words;
using AsposeMcpServer.Core;
using SaveFormat = Aspose.Words.SaveFormat;

namespace AsposeMcpServer.Tools.Conversion;

/// <summary>
///     Tool for converting documents (Word, Excel, PowerPoint) to PDF format
/// </summary>
public class ConvertToPdfTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples
    /// </summary>
    public string Description => @"Convert any document (Word, Excel, PowerPoint) to PDF.

Usage examples:
- Convert Word to PDF: convert_to_pdf(inputPath='doc.docx', outputPath='doc.pdf')
- Convert Excel to PDF: convert_to_pdf(inputPath='book.xlsx', outputPath='book.pdf')
- Convert PowerPoint to PDF: convert_to_pdf(inputPath='presentation.pptx', outputPath='presentation.pdf')";

    /// <summary>
    ///     Gets the JSON schema defining the input parameters for the tool
    /// </summary>
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

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var inputPath = ArgumentHelper.GetString(arguments, "inputPath");
        var outputPath = ArgumentHelper.GetString(arguments, "outputPath");

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
                wordDoc.Save(outputPath, SaveFormat.Pdf);
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