using System.ComponentModel;
using Aspose.Cells;
using Aspose.Slides;
using Aspose.Words;
using AsposeMcpServer.Core.Helpers;
using ModelContextProtocol.Server;
using SaveFormat = Aspose.Words.SaveFormat;

namespace AsposeMcpServer.Tools.Conversion;

/// <summary>
///     Tool for converting documents between various formats with automatic source type detection
///     Supports Word, Excel, and PowerPoint documents
/// </summary>
[McpServerToolType]
public class ConvertDocumentTool
{
    [McpServerTool(Name = "convert_document")]
    [Description(@"Convert documents between various formats (auto-detect source type).

Usage examples:
- Convert Word to HTML: convert_document(inputPath='doc.docx', outputPath='doc.html')
- Convert Excel to CSV: convert_document(inputPath='book.xlsx', outputPath='book.csv')
- Convert PowerPoint to PDF: convert_document(inputPath='presentation.pptx', outputPath='presentation.pdf')
- Convert PDF to Word: convert_document(inputPath='doc.pdf', outputPath='doc.docx')")]
    public string Execute(
        [Description("Input file path (required, auto-detects format from extension)")]
        string inputPath,
        [Description("Output file path (required, format determined by extension)")]
        string outputPath)
    {
        SecurityHelper.ValidateFilePath(inputPath, "inputPath", true);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        var inputExtension = Path.GetExtension(inputPath).ToLower();
        var outputExtension = Path.GetExtension(outputPath).ToLower();

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

        return $"Document converted from {inputExtension} to {outputExtension} format. Output: {outputPath}";
    }

    /// <summary>
    ///     Determines whether the specified file extension corresponds to a Word document format.
    /// </summary>
    /// <param name="extension">The file extension to check (including the leading dot).</param>
    /// <returns><c>true</c> if the extension is a Word document format; otherwise, <c>false</c>.</returns>
    private static bool IsWordDocument(string extension)
    {
        return extension is ".doc" or ".docx" or ".rtf" or ".odt" or ".txt";
    }

    /// <summary>
    ///     Determines whether the specified file extension corresponds to an Excel document format.
    /// </summary>
    /// <param name="extension">The file extension to check (including the leading dot).</param>
    /// <returns><c>true</c> if the extension is an Excel document format; otherwise, <c>false</c>.</returns>
    private static bool IsExcelDocument(string extension)
    {
        return extension is ".xls" or ".xlsx" or ".csv" or ".ods";
    }

    /// <summary>
    ///     Determines whether the specified file extension corresponds to a PowerPoint presentation format.
    /// </summary>
    /// <param name="extension">The file extension to check (including the leading dot).</param>
    /// <returns><c>true</c> if the extension is a PowerPoint presentation format; otherwise, <c>false</c>.</returns>
    private static bool IsPresentationDocument(string extension)
    {
        return extension is ".ppt" or ".pptx" or ".odp";
    }

    /// <summary>
    ///     Gets the Aspose.Words save format corresponding to the specified file extension.
    /// </summary>
    /// <param name="extension">The target file extension (including the leading dot).</param>
    /// <returns>The <see cref="SaveFormat" /> value for the specified extension.</returns>
    /// <exception cref="ArgumentException">Thrown when the extension is not a supported output format.</exception>
    private static SaveFormat GetWordSaveFormat(string extension)
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

    /// <summary>
    ///     Gets the Aspose.Cells save format corresponding to the specified file extension.
    /// </summary>
    /// <param name="extension">The target file extension (including the leading dot).</param>
    /// <returns>The <see cref="Aspose.Cells.SaveFormat" /> value for the specified extension.</returns>
    /// <exception cref="ArgumentException">Thrown when the extension is not a supported output format.</exception>
    private static Aspose.Cells.SaveFormat GetExcelSaveFormat(string extension)
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

    /// <summary>
    ///     Gets the Aspose.Slides save format corresponding to the specified file extension.
    /// </summary>
    /// <param name="extension">The target file extension (including the leading dot).</param>
    /// <returns>The <see cref="Aspose.Slides.Export.SaveFormat" /> value for the specified extension.</returns>
    /// <exception cref="ArgumentException">Thrown when the extension is not a supported output format.</exception>
    private static Aspose.Slides.Export.SaveFormat GetPresentationSaveFormat(string extension)
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