using System.Text.Json.Nodes;
using Aspose.Cells;
using Aspose.Slides;
using Aspose.Words;
using AsposeMcpServer.Core;
using SaveFormat = Aspose.Words.SaveFormat;

namespace AsposeMcpServer.Tools.Conversion;

/// <summary>
///     Tool for converting documents between various formats with automatic source type detection
///     Supports Word, Excel, and PowerPoint documents
/// </summary>
public class ConvertDocumentTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples
    /// </summary>
    public string Description => @"Convert documents between various formats (auto-detect source type).

Usage examples:
- Convert Word to HTML: convert_document(inputPath='doc.docx', outputPath='doc.html')
- Convert Excel to CSV: convert_document(inputPath='book.xlsx', outputPath='book.csv')
- Convert PowerPoint to PDF: convert_document(inputPath='presentation.pptx', outputPath='presentation.pdf')
- Convert PDF to Word: convert_document(inputPath='doc.pdf', outputPath='doc.docx')";

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
                description = "Input file path (required, auto-detects format from extension)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (required, format determined by extension)"
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

        var inputExtension = Path.GetExtension(inputPath).ToLower();
        var outputExtension = Path.GetExtension(outputPath).ToLower();

        // Detect document type and convert
        if (IsWordDocument(inputExtension))
        {
            var doc = new Document(inputPath);
            var saveFormat = GetWordSaveFormat(outputExtension);
            doc.Save(outputPath, saveFormat);
        }
        else if (IsExcelDocument(inputExtension))
        {
            using var workbook = new Workbook(inputPath);
            var saveFormat = GetExcelSaveFormat(outputExtension);
            workbook.Save(outputPath, saveFormat);
        }
        else if (IsPresentationDocument(inputExtension))
        {
            using var presentation = new Presentation(inputPath);
            var saveFormat = GetPresentationSaveFormat(outputExtension);
            presentation.Save(outputPath, saveFormat);
        }
        else
        {
            throw new ArgumentException($"Unsupported input format: {inputExtension}");
        }

        return await Task.FromResult($"Document converted from {inputPath} to {outputPath}");
    }

    private bool IsWordDocument(string extension)
    {
        return extension is ".doc" or ".docx" or ".rtf" or ".odt" or ".txt";
    }

    private bool IsExcelDocument(string extension)
    {
        return extension is ".xls" or ".xlsx" or ".csv" or ".ods";
    }

    private bool IsPresentationDocument(string extension)
    {
        return extension is ".ppt" or ".pptx" or ".odp";
    }

    private SaveFormat GetWordSaveFormat(string extension)
    {
        return extension switch
        {
            ".pdf" => SaveFormat.Pdf,
            ".docx" => SaveFormat.Docx,
            ".doc" => SaveFormat.Doc,
            ".rtf" => SaveFormat.Rtf,
            ".html" => SaveFormat.Html,
            ".txt" => SaveFormat.Text,
            ".odt" => SaveFormat.Odt,
            _ => throw new ArgumentException($"Unsupported output format: {extension}")
        };
    }

    private Aspose.Cells.SaveFormat GetExcelSaveFormat(string extension)
    {
        return extension switch
        {
            ".pdf" => Aspose.Cells.SaveFormat.Pdf,
            ".xlsx" => Aspose.Cells.SaveFormat.Xlsx,
            ".xls" => Aspose.Cells.SaveFormat.Excel97To2003,
            ".csv" => Aspose.Cells.SaveFormat.Csv,
            ".html" => Aspose.Cells.SaveFormat.Html,
            ".ods" => Aspose.Cells.SaveFormat.Ods,
            _ => throw new ArgumentException($"Unsupported output format: {extension}")
        };
    }

    private Aspose.Slides.Export.SaveFormat GetPresentationSaveFormat(string extension)
    {
        return extension switch
        {
            ".pdf" => Aspose.Slides.Export.SaveFormat.Pdf,
            ".pptx" => Aspose.Slides.Export.SaveFormat.Pptx,
            ".ppt" => Aspose.Slides.Export.SaveFormat.Ppt,
            ".html" => Aspose.Slides.Export.SaveFormat.Html,
            ".odp" => Aspose.Slides.Export.SaveFormat.Odp,
            _ => throw new ArgumentException($"Unsupported output format: {extension}")
        };
    }
}