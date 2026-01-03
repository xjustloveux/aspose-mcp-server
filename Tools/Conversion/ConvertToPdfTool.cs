using System.ComponentModel;
using Aspose.Cells;
using Aspose.Slides;
using Aspose.Words;
using AsposeMcpServer.Core.Helpers;
using ModelContextProtocol.Server;
using SaveFormat = Aspose.Words.SaveFormat;

namespace AsposeMcpServer.Tools.Conversion;

/// <summary>
///     Tool for converting documents (Word, Excel, PowerPoint) to PDF format
/// </summary>
[McpServerToolType]
public class ConvertToPdfTool
{
    [McpServerTool(Name = "convert_to_pdf")]
    [Description(@"Convert any document (Word, Excel, PowerPoint) to PDF.

Usage examples:
- Convert Word to PDF: convert_to_pdf(inputPath='doc.docx', outputPath='doc.pdf')
- Convert Excel to PDF: convert_to_pdf(inputPath='book.xlsx', outputPath='book.pdf')
- Convert PowerPoint to PDF: convert_to_pdf(inputPath='presentation.pptx', outputPath='presentation.pdf')")]
    public string Execute(
        [Description("Input file path (required, supports Word, Excel, PowerPoint formats)")]
        string inputPath,
        [Description("Output PDF file path (required)")]
        string outputPath)
    {
        SecurityHelper.ValidateFilePath(inputPath, "inputPath", true);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

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

        return $"Document converted to PDF. Output: {outputPath}";
    }
}