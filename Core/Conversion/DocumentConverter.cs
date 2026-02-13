using System.Drawing.Imaging;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;
using Aspose.Pdf;
using Aspose.Pdf.Devices;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core.Progress;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol;
using TxtSaveOptions = Aspose.Cells.TxtSaveOptions;
using Document = Aspose.Words.Document;
using HtmlLoadOptions = Aspose.Pdf.HtmlLoadOptions;
using HtmlSaveOptions = Aspose.Cells.HtmlSaveOptions;
using WordHtmlSaveOptions = Aspose.Words.Saving.HtmlSaveOptions;
using ImageSaveOptions = Aspose.Words.Saving.ImageSaveOptions;
using PageSet = Aspose.Words.Saving.PageSet;
using PdfCompliance = Aspose.Words.Saving.PdfCompliance;
using PdfSaveOptions = Aspose.Words.Saving.PdfSaveOptions;
using SaveFormat = Aspose.Cells.SaveFormat;
using SvgSaveOptions = Aspose.Pdf.SvgSaveOptions;
using WordSaveFormat = Aspose.Words.SaveFormat;

namespace AsposeMcpServer.Core.Conversion;

/// <summary>
///     Options for document conversion operations.
/// </summary>
public class ConversionOptions
{
    /// <summary>
    ///     Optional 1-based page/sheet index for single page/sheet output.
    /// </summary>
    public int? PageIndex { get; init; }

    /// <summary>
    ///     Resolution in DPI for image output. Default is 150.
    /// </summary>
    public int Dpi { get; init; } = 150;

    /// <summary>
    ///     Whether to embed images as Base64 in HTML output. Default is true.
    /// </summary>
    public bool HtmlEmbedImages { get; init; } = true;

    /// <summary>
    ///     Whether to export as single HTML file without external resources. Default is true.
    /// </summary>
    public bool HtmlSingleFile { get; init; } = true;

    /// <summary>
    ///     JPEG quality (1-100) for JPEG image output. Default is 90.
    /// </summary>
    public int JpegQuality { get; init; } = 90;

    /// <summary>
    ///     CSV field separator character. Default is comma.
    /// </summary>
    public string CsvSeparator { get; init; } = ",";

    /// <summary>
    ///     PDF/A compliance level for Word documents: "PDFA1A", "PDFA1B", "PDFA2A", "PDFA2U", "PDFA4".
    ///     Null means no specific compliance.
    /// </summary>
    public string? PdfCompliance { get; init; }
}

/// <summary>
///     Provides document conversion functionality for various Aspose document types.
///     This utility class is shared between Extension system and Conversion tools.
/// </summary>
public static class DocumentConverter
{
    /// <summary>
    ///     MIME type mappings for output formats.
    /// </summary>
    private static readonly Dictionary<string, string> MimeTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        { "pdf", "application/pdf" },
        { "html", "text/html" },
        { "htm", "text/html" },
        { "png", "image/png" },
        { "jpg", "image/jpeg" },
        { "jpeg", "image/jpeg" },
        { "tiff", "image/tiff" },
        { "tif", "image/tiff" },
        { "docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
        { "doc", "application/msword" },
        { "rtf", "application/rtf" },
        { "txt", "text/plain" },
        { "odt", "application/vnd.oasis.opendocument.text" },
        { "xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
        { "xls", "application/vnd.ms-excel" },
        { "csv", "text/csv" },
        { "ods", "application/vnd.oasis.opendocument.spreadsheet" },
        { "pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation" },
        { "ppt", "application/vnd.ms-powerpoint" },
        { "odp", "application/vnd.oasis.opendocument.presentation" },
        { "epub", "application/epub+zip" },
        { "svg", "image/svg+xml" },
        { "xps", "application/vnd.ms-xpsdocument" },
        { "xml", "application/xml" },
        { "md", "text/markdown" },
        { "tex", "application/x-tex" },
        { "mht", "message/rfc822" },
        { "mhtml", "message/rfc822" }
    };

    /// <summary>
    ///     Supported output formats for each document type.
    /// </summary>
    private static readonly Dictionary<DocumentType, HashSet<string>> SupportedFormats =
        new()
        {
            {
                DocumentType.Word,
                new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "pdf", "html", "docx", "doc", "rtf", "txt", "odt", "png", "jpg", "jpeg", "tiff", "tif", "bmp", "svg"
                }
            },
            {
                DocumentType.Excel,
                new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                    { "pdf", "html", "xlsx", "xls", "csv", "ods", "png", "jpg", "jpeg", "tiff", "tif", "bmp", "svg" }
            },
            {
                DocumentType.PowerPoint,
                new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                    { "pdf", "html", "pptx", "ppt", "odp", "png", "jpg" }
            },
            {
                DocumentType.Pdf,
                new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "docx", "doc", "html", "xlsx", "pptx", "png", "jpg", "jpeg", "tiff", "tif", "epub", "svg", "xps",
                    "xml"
                }
            }
        };

    #region Document Type Detection

    /// <summary>
    ///     Determines whether the specified extension represents a Word document.
    /// </summary>
    /// <param name="extension">The file extension to check (with or without leading dot).</param>
    /// <returns><c>true</c> if the extension is a Word document format; otherwise, <c>false</c>.</returns>
    public static bool IsWordDocument(string extension)
    {
        var ext = NormalizeExtension(extension);
        return ext is "doc" or "docx" or "rtf" or "odt" or "txt";
    }

    /// <summary>
    ///     Determines whether the specified extension represents an Excel document.
    /// </summary>
    /// <param name="extension">The file extension to check (with or without leading dot).</param>
    /// <returns><c>true</c> if the extension is an Excel document format; otherwise, <c>false</c>.</returns>
    public static bool IsExcelDocument(string extension)
    {
        var ext = NormalizeExtension(extension);
        return ext is "xls" or "xlsx" or "csv" or "ods";
    }

    /// <summary>
    ///     Determines whether the specified extension represents a PowerPoint presentation.
    /// </summary>
    /// <param name="extension">The file extension to check (with or without leading dot).</param>
    /// <returns><c>true</c> if the extension is a PowerPoint format; otherwise, <c>false</c>.</returns>
    public static bool IsPowerPointDocument(string extension)
    {
        var ext = NormalizeExtension(extension);
        return ext is "ppt" or "pptx" or "odp";
    }

    /// <summary>
    ///     Determines whether the specified extension represents a PDF document.
    /// </summary>
    /// <param name="extension">The file extension to check (with or without leading dot).</param>
    /// <returns><c>true</c> if the extension is a PDF document format; otherwise, <c>false</c>.</returns>
    public static bool IsPdfDocument(string extension)
    {
        var ext = NormalizeExtension(extension);
        return ext is "pdf";
    }

    /// <summary>
    ///     Determines whether the specified extension represents an image format.
    /// </summary>
    /// <param name="extension">The file extension to check (with or without leading dot).</param>
    /// <returns><c>true</c> if the extension is an image format; otherwise, <c>false</c>.</returns>
    public static bool IsImageFormat(string extension)
    {
        var ext = NormalizeExtension(extension);
        return ext is "png" or "jpg" or "jpeg" or "tiff" or "tif";
    }

    /// <summary>
    ///     Determines whether the specified format is an image format for Excel conversion.
    /// </summary>
    /// <param name="format">The normalized format string.</param>
    /// <returns><c>true</c> if the format is an image format for Excel; otherwise, <c>false</c>.</returns>
    public static bool IsExcelImageFormat(string format)
    {
        var ext = NormalizeExtension(format);
        return ext is "png" or "jpg" or "jpeg" or "tiff" or "tif" or "bmp" or "svg";
    }

    /// <summary>
    ///     Determines whether the specified extension represents a format that can be converted to PDF.
    ///     Includes HTML, EPUB, Markdown, SVG, XPS, LaTeX, and MHT formats.
    /// </summary>
    /// <param name="extension">The file extension to check (with or without leading dot).</param>
    /// <returns><c>true</c> if the extension can be converted to PDF; otherwise, <c>false</c>.</returns>
    public static bool IsPdfConvertibleFormat(string extension)
    {
        var ext = NormalizeExtension(extension);
        return ext is "html" or "htm" or "epub" or "md" or "svg" or "xps" or "tex" or "mht" or "mhtml";
    }

    /// <summary>
    ///     Gets the document type for the specified file extension.
    /// </summary>
    /// <param name="extension">The file extension to check (with or without leading dot).</param>
    /// <returns>The document type, or <c>null</c> if the extension is not recognized.</returns>
    public static DocumentType? GetDocumentType(string extension)
    {
        if (IsWordDocument(extension)) return DocumentType.Word;
        if (IsExcelDocument(extension)) return DocumentType.Excel;
        if (IsPowerPointDocument(extension)) return DocumentType.PowerPoint;
        if (IsPdfDocument(extension)) return DocumentType.Pdf;
        return null;
    }

    #endregion

    #region Save Format Helpers

    /// <summary>
    ///     Gets the Word save format for the specified format string.
    /// </summary>
    /// <param name="format">The target format (with or without leading dot).</param>
    /// <returns>The corresponding Word save format.</returns>
    /// <exception cref="ArgumentException">Thrown when the format is not supported.</exception>
    public static WordSaveFormat GetWordSaveFormat(string format)
    {
        var ext = NormalizeExtension(format);
        return ext switch
        {
            "pdf" => WordSaveFormat.Pdf,
            "docx" => WordSaveFormat.Docx,
            "doc" => WordSaveFormat.Doc,
            "rtf" => WordSaveFormat.Rtf,
            "html" => WordSaveFormat.Html,
            "txt" => WordSaveFormat.Text,
            "odt" => WordSaveFormat.Odt,
            _ => throw new ArgumentException($"Unsupported output format for Word: {format}")
        };
    }

    /// <summary>
    ///     Gets the Excel save format for the specified format string.
    /// </summary>
    /// <param name="format">The target format (with or without leading dot).</param>
    /// <returns>The corresponding Excel save format.</returns>
    /// <exception cref="ArgumentException">Thrown when the format is not supported.</exception>
    public static SaveFormat GetExcelSaveFormat(string format)
    {
        var ext = NormalizeExtension(format);
        return ext switch
        {
            "pdf" => SaveFormat.Pdf,
            "xlsx" => SaveFormat.Xlsx,
            "xls" => SaveFormat.Excel97To2003,
            "csv" => SaveFormat.Csv,
            "html" => SaveFormat.Html,
            "ods" => SaveFormat.Ods,
            _ => throw new ArgumentException($"Unsupported output format for Excel: {format}")
        };
    }

    /// <summary>
    ///     Gets the PowerPoint save format for the specified format string.
    /// </summary>
    /// <param name="format">The target format (with or without leading dot).</param>
    /// <returns>The corresponding PowerPoint save format.</returns>
    /// <exception cref="ArgumentException">Thrown when the format is not supported.</exception>
    public static Aspose.Slides.Export.SaveFormat GetPresentationSaveFormat(string format)
    {
        var ext = NormalizeExtension(format);
        return ext switch
        {
            "pdf" => Aspose.Slides.Export.SaveFormat.Pdf,
            "pptx" => Aspose.Slides.Export.SaveFormat.Pptx,
            "ppt" => Aspose.Slides.Export.SaveFormat.Ppt,
            "html" => Aspose.Slides.Export.SaveFormat.Html,
            "odp" => Aspose.Slides.Export.SaveFormat.Odp,
            _ => throw new ArgumentException($"Unsupported output format for PowerPoint: {format}")
        };
    }

    /// <summary>
    ///     Gets the PDF save format for the specified format string.
    /// </summary>
    /// <param name="format">The target format (with or without leading dot).</param>
    /// <returns>The corresponding PDF save format.</returns>
    /// <exception cref="ArgumentException">Thrown when the format is not supported.</exception>
    public static Aspose.Pdf.SaveFormat GetPdfSaveFormat(string format)
    {
        var ext = NormalizeExtension(format);
        return ext switch
        {
            "docx" => Aspose.Pdf.SaveFormat.DocX,
            "doc" => Aspose.Pdf.SaveFormat.Doc,
            "html" => Aspose.Pdf.SaveFormat.Html,
            "xlsx" => Aspose.Pdf.SaveFormat.Excel,
            "pptx" => Aspose.Pdf.SaveFormat.Pptx,
            "txt" => Aspose.Pdf.SaveFormat.TeX,
            "epub" => Aspose.Pdf.SaveFormat.Epub,
            "svg" => Aspose.Pdf.SaveFormat.Svg,
            "xps" => Aspose.Pdf.SaveFormat.Xps,
            "xml" => Aspose.Pdf.SaveFormat.Xml,
            _ => throw new ArgumentException($"Unsupported output format for PDF: {format}")
        };
    }

    #endregion

    #region Stream Conversion (for Extension system)

    /// <summary>
    ///     Converts a document to the specified format and returns a Stream.
    /// </summary>
    /// <param name="document">The Aspose document object (Document, Workbook, Presentation, or Aspose.Pdf.Document).</param>
    /// <param name="documentType">The type of the source document.</param>
    /// <param name="outputFormat">The target output format (e.g., "pdf", "html", "png").</param>
    /// <param name="options">Optional conversion options. If null, defaults are used.</param>
    /// <returns>A MemoryStream containing the converted document.</returns>
    /// <exception cref="ArgumentNullException">Thrown when document is null.</exception>
    /// <exception cref="ArgumentException">Thrown when the output format is not supported for the document type.</exception>
    public static Stream ConvertToStream(object document, DocumentType documentType, string outputFormat,
        ConversionOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(document);

        if (string.IsNullOrWhiteSpace(outputFormat))
            throw new ArgumentException("Output format cannot be null or empty.", nameof(outputFormat));

        var format = NormalizeExtension(outputFormat);

        if (!IsFormatSupported(documentType, format))
            throw new ArgumentException(
                $"Output format '{format}' is not supported for document type '{documentType}'.");

        var stream = new MemoryStream();

        try
        {
            ConvertToStreamInternal(document, documentType, format, stream, null, options);
            stream.Position = 0;
            return stream;
        }
        catch
        {
            stream.Dispose();
            throw;
        }
    }

    /// <summary>
    ///     Converts a document to the specified format and returns a byte array.
    /// </summary>
    /// <param name="document">The Aspose document object.</param>
    /// <param name="documentType">The type of the source document.</param>
    /// <param name="outputFormat">The target output format.</param>
    /// <param name="options">Optional conversion options. If null, defaults are used.</param>
    /// <returns>A byte array containing the converted document.</returns>
    /// <exception cref="ArgumentNullException">Thrown when document is null.</exception>
    /// <exception cref="ArgumentException">Thrown when the output format is not supported for the document type.</exception>
    public static byte[] ConvertToBytes(object document, DocumentType documentType, string outputFormat,
        ConversionOptions? options = null)
    {
        using var stream = ConvertToStream(document, documentType, outputFormat, options);
        return ((MemoryStream)stream).ToArray();
    }

    #endregion

    #region File Conversion (for Tools)

    /// <summary>
    ///     Converts a Word document to the specified output format.
    /// </summary>
    /// <param name="document">The Word document to convert.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="outputFormat">The target output format (with or without leading dot).</param>
    /// <param name="progress">Optional progress reporter (only effective for PDF output).</param>
    /// <param name="options">Optional conversion options. If null, defaults are used.</param>
    /// <exception cref="ArgumentException">Thrown when the output format is not supported.</exception>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when pageIndex is out of range.</exception>
    public static void ConvertWordDocument(Document document, string outputPath, string outputFormat,
        IProgress<ProgressNotificationValue>? progress = null, ConversionOptions? options = null)
    {
        options ??= new ConversionOptions();
        document.UpdatePageLayout();
        var format = NormalizeExtension(outputFormat);

        if (IsWordImageFormat(format))
        {
            ConvertWordToImages(document, outputPath, format, options);
            return;
        }

        if (format == "pdf")
        {
            var saveOptions = new PdfSaveOptions
            {
                ProgressCallback = new WordsProgressAdapter(progress)
            };
            ApplyWordPdfCompliance(saveOptions, options.PdfCompliance);
            document.Save(outputPath, saveOptions);
        }
        else if (format is "html" or "htm")
        {
            var saveOptions = new WordHtmlSaveOptions
            {
                ExportImagesAsBase64 = options.HtmlEmbedImages || options.HtmlSingleFile,
                ExportFontsAsBase64 = options.HtmlSingleFile
            };
            document.Save(outputPath, saveOptions);
        }
        else
        {
            var saveFormat = GetWordSaveFormat(format);
            document.Save(outputPath, saveFormat);
        }
    }

    /// <summary>
    ///     Applies PDF compliance settings to Word PDF save options.
    /// </summary>
    /// <param name="saveOptions">The PDF save options to configure.</param>
    /// <param name="compliance">The compliance string (e.g., "PDFA1A", "PDFA1B").</param>
    private static void ApplyWordPdfCompliance(PdfSaveOptions saveOptions, string? compliance)
    {
        if (string.IsNullOrEmpty(compliance))
            return;

        saveOptions.Compliance = compliance.ToUpperInvariant() switch
        {
            "PDFA1A" => PdfCompliance.PdfA1a,
            "PDFA1B" => PdfCompliance.PdfA1b,
            "PDFA2A" => PdfCompliance.PdfA2a,
            "PDFA2U" => PdfCompliance.PdfA2u,
            "PDFA4" => PdfCompliance.PdfA4,
            _ => saveOptions.Compliance
        };
    }

    /// <summary>
    ///     Determines whether the specified format is an image format for Word conversion.
    /// </summary>
    /// <param name="format">The normalized format string.</param>
    /// <returns><c>true</c> if the format is an image format for Word; otherwise, <c>false</c>.</returns>
    private static bool IsWordImageFormat(string format)
    {
        return format is "png" or "jpg" or "jpeg" or "tiff" or "tif" or "bmp" or "svg";
    }

    /// <summary>
    ///     Converts an Excel workbook to the specified output format.
    /// </summary>
    /// <param name="workbook">The Excel workbook to convert.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="outputFormat">The target output format (with or without leading dot).</param>
    /// <param name="progress">Optional progress reporter (only effective for PDF output).</param>
    /// <param name="options">Optional conversion options. If null, defaults are used.</param>
    /// <exception cref="ArgumentException">Thrown when the output format is not supported.</exception>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when sheetIndex is out of range.</exception>
    public static void ConvertExcelDocument(Workbook workbook, string outputPath, string outputFormat,
        IProgress<ProgressNotificationValue>? progress = null, ConversionOptions? options = null)
    {
        options ??= new ConversionOptions();
        workbook.CalculateFormula();
        var format = NormalizeExtension(outputFormat);

        if (IsExcelImageFormat(format))
        {
            ConvertExcelToImages(workbook, outputPath, format, options);
            return;
        }

        if (format == "pdf")
        {
            var saveOptions = new Aspose.Cells.PdfSaveOptions
            {
                PageSavingCallback = new CellsProgressAdapter(progress)
            };
            ApplyExcelPdfCompliance(saveOptions, options.PdfCompliance);
            workbook.Save(outputPath, saveOptions);
        }
        else if (format is "html" or "htm")
        {
            var saveOptions = new HtmlSaveOptions
            {
                ExportImagesAsBase64 = options.HtmlEmbedImages || options.HtmlSingleFile,
                SaveAsSingleFile = options.HtmlSingleFile,
                ShowAllSheets = options.HtmlSingleFile
            };

            workbook.Save(outputPath, saveOptions);
        }
        else if (format == "csv")
        {
            var saveOptions = new TxtSaveOptions
            {
                Separator = string.IsNullOrEmpty(options.CsvSeparator) ? ',' : options.CsvSeparator[0]
            };
            workbook.Save(outputPath, saveOptions);
        }
        else
        {
            var saveFormat = GetExcelSaveFormat(format);
            workbook.Save(outputPath, saveFormat);
        }
    }

    /// <summary>
    ///     Applies PDF compliance settings to Excel PDF save options.
    /// </summary>
    /// <param name="saveOptions">The PDF save options to configure.</param>
    /// <param name="compliance">The compliance string (e.g., "PDFA1A", "PDFA1B").</param>
    private static void ApplyExcelPdfCompliance(Aspose.Cells.PdfSaveOptions saveOptions, string? compliance)
    {
        if (string.IsNullOrEmpty(compliance))
            return;

        saveOptions.Compliance = compliance.ToUpperInvariant() switch
        {
            "PDFA1A" => Aspose.Cells.Rendering.PdfCompliance.PdfA1a,
            "PDFA1B" => Aspose.Cells.Rendering.PdfCompliance.PdfA1b,
            _ => saveOptions.Compliance
        };
    }

    /// <summary>
    ///     Converts a PowerPoint presentation to the specified output format.
    /// </summary>
    /// <param name="presentation">The PowerPoint presentation to convert.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="outputFormat">The target output format (with or without leading dot).</param>
    /// <param name="progress">Optional progress reporter (only effective for PDF output).</param>
    /// <param name="options">Optional conversion options. If null, defaults are used.</param>
    /// <exception cref="ArgumentException">Thrown when the output format is not supported.</exception>
    public static void ConvertPowerPointDocument(Presentation presentation, string outputPath, string outputFormat,
        IProgress<ProgressNotificationValue>? progress = null, ConversionOptions? options = null)
    {
        options ??= new ConversionOptions();
        var format = NormalizeExtension(outputFormat);

        if (format == "pdf")
        {
            var saveOptions = new PdfOptions
            {
                ProgressCallback = new SlidesProgressAdapter(progress),
                JpegQuality = (byte)options.JpegQuality
            };
            ApplySlidesPdfCompliance(saveOptions, options.PdfCompliance);
            presentation.Save(outputPath, Aspose.Slides.Export.SaveFormat.Pdf, saveOptions);
        }
        else
        {
            var saveFormat = GetPresentationSaveFormat(format);
            presentation.Save(outputPath, saveFormat);
        }
    }

    /// <summary>
    ///     Applies PDF compliance settings to PowerPoint PDF save options.
    /// </summary>
    /// <param name="saveOptions">The PDF save options to configure.</param>
    /// <param name="compliance">The compliance string (e.g., "PDFA1A", "PDFA1B", "PDFUA").</param>
    private static void ApplySlidesPdfCompliance(PdfOptions saveOptions, string? compliance)
    {
        if (string.IsNullOrEmpty(compliance))
            return;

        saveOptions.Compliance = compliance.ToUpperInvariant() switch
        {
            "PDFA1A" => Aspose.Slides.Export.PdfCompliance.PdfA1a,
            "PDFA1B" => Aspose.Slides.Export.PdfCompliance.PdfA1b,
            "PDFA2A" => Aspose.Slides.Export.PdfCompliance.PdfA2a,
            "PDFA2B" => Aspose.Slides.Export.PdfCompliance.PdfA2b,
            "PDFA3A" => Aspose.Slides.Export.PdfCompliance.PdfA3a,
            "PDFA3B" => Aspose.Slides.Export.PdfCompliance.PdfA3b,
            "PDFUA" => Aspose.Slides.Export.PdfCompliance.PdfUa,
            _ => saveOptions.Compliance
        };
    }

    /// <summary>
    ///     Converts a PDF document to the specified output format.
    ///     For image formats, converts all pages (PNG/JPEG: one file per page, TIFF: single multi-page file).
    /// </summary>
    /// <param name="pdfDocument">The PDF document to convert.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="outputFormat">The target output format (with or without leading dot).</param>
    /// <param name="options">Optional conversion options. If null, defaults are used.</param>
    /// <exception cref="ArgumentException">Thrown when the output format is not supported.</exception>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when pageIndex is out of range.</exception>
    public static void ConvertPdfDocument(Aspose.Pdf.Document pdfDocument, string outputPath, string outputFormat,
        ConversionOptions? options = null)
    {
        options ??= new ConversionOptions();
        var format = NormalizeExtension(outputFormat);

        if (IsImageFormat(format))
        {
            ConvertPdfToImages(pdfDocument, outputPath, format, options.PageIndex);
        }
        else if (format is "html" or "htm")
        {
            var htmlOptions = new Aspose.Pdf.HtmlSaveOptions
            {
                PartsEmbeddingMode = Aspose.Pdf.HtmlSaveOptions.PartsEmbeddingModes.EmbedAllIntoHtml,
                RasterImagesSavingMode =
                    Aspose.Pdf.HtmlSaveOptions.RasterImagesSavingModes.AsEmbeddedPartsOfPngPageBackground,
                FontSavingMode = Aspose.Pdf.HtmlSaveOptions.FontSavingModes.SaveInAllFormats
            };
            pdfDocument.Save(outputPath, htmlOptions);
        }
        else if (format == "svg")
        {
            var svgOptions = new SvgSaveOptions
            {
                CompressOutputToZipArchive = false
            };
            pdfDocument.Save(outputPath, svgOptions);
        }
        else
        {
            var saveFormat = GetPdfSaveFormat(format);
            pdfDocument.Save(outputPath, saveFormat);
        }
    }

    /// <summary>
    ///     Converts a PDF file to images, one file per page (PNG/JPEG) or a single multi-page file (TIFF).
    /// </summary>
    /// <param name="inputPath">The input PDF file path.</param>
    /// <param name="outputPath">The output image file path (page number will be appended for multi-page PNG/JPEG).</param>
    /// <param name="outputFormat">The target image format (with or without leading dot).</param>
    /// <param name="pageIndex">Optional 1-based page index for single page output (omit for all pages).</param>
    /// <param name="options">Optional conversion options. If null, defaults are used.</param>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when pageIndex is out of range.</exception>
    /// <exception cref="ArgumentException">Thrown when the image format is not supported.</exception>
    public static void ConvertPdfToImages(string inputPath, string outputPath, string outputFormat,
        int? pageIndex = null, ConversionOptions? options = null)
    {
        using var pdfDoc = new Aspose.Pdf.Document(inputPath);
        ConvertPdfToImages(pdfDoc, outputPath, outputFormat, pageIndex, options);
    }

    /// <summary>
    ///     Converts a PDF document to images, one file per page (PNG/JPEG) or a single multi-page file (TIFF).
    /// </summary>
    /// <param name="pdfDocument">The PDF document to convert.</param>
    /// <param name="outputPath">The output image file path (page number will be appended for multi-page PNG/JPEG).</param>
    /// <param name="outputFormat">The target image format (with or without leading dot).</param>
    /// <param name="pageIndex">Optional 1-based page index for single page output (omit for all pages).</param>
    /// <param name="options">Optional conversion options. If null, defaults are used.</param>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when pageIndex is out of range.</exception>
    /// <exception cref="ArgumentException">Thrown when the image format is not supported.</exception>
    public static void ConvertPdfToImages(Aspose.Pdf.Document pdfDocument, string outputPath, string outputFormat,
        int? pageIndex = null, ConversionOptions? options = null)
    {
        options ??= new ConversionOptions();
        var format = NormalizeExtension(outputFormat);
        var resolution = new Resolution(options.Dpi);

        if (pageIndex.HasValue)
        {
            if (pageIndex.Value < 1 || pageIndex.Value > pdfDocument.Pages.Count)
                throw new ArgumentOutOfRangeException(nameof(pageIndex),
                    $"Page index must be between 1 and {pdfDocument.Pages.Count}");

            if (format is "tiff" or "tif")
            {
                var tiffDevice = new TiffDevice(resolution);
                using var stream = new FileStream(outputPath, FileMode.Create);
                tiffDevice.Process(pdfDocument, pageIndex.Value, pageIndex.Value, stream);
            }
            else
            {
                using var stream = new FileStream(outputPath, FileMode.Create);
                PageDevice device = format switch
                {
                    "png" => new PngDevice(resolution),
                    "jpg" or "jpeg" => new JpegDevice(resolution, options.JpegQuality),
                    _ => throw new ArgumentException($"Unsupported image format: {format}")
                };

                device.Process(pdfDocument.Pages[pageIndex.Value], stream);
            }

            return;
        }

        if (format is "tiff" or "tif")
        {
            var tiffDevice = new TiffDevice(resolution);
            using var stream = new FileStream(outputPath, FileMode.Create);
            tiffDevice.Process(pdfDocument, stream);
            return;
        }

        var dir = Path.GetDirectoryName(outputPath) ?? ".";
        var nameWithoutExt = Path.GetFileNameWithoutExtension(outputPath);
        var ext = Path.GetExtension(outputPath);

        for (var i = 1; i <= pdfDocument.Pages.Count; i++)
        {
            var pagePath = pdfDocument.Pages.Count == 1
                ? outputPath
                : Path.Combine(dir, $"{nameWithoutExt}_{i}{ext}");

            using var stream = new FileStream(pagePath, FileMode.Create);
            PageDevice device = format switch
            {
                "png" => new PngDevice(resolution),
                "jpg" or "jpeg" => new JpegDevice(resolution, options.JpegQuality),
                _ => throw new ArgumentException($"Unsupported image format: {format}")
            };

            device.Process(pdfDocument.Pages[i], stream);
        }
    }

    /// <summary>
    ///     Converts Word document pages to images.
    /// </summary>
    /// <param name="document">The Word document to convert.</param>
    /// <param name="outputPath">The output file path (page number will be appended for multi-page output).</param>
    /// <param name="outputFormat">The target image format (png, jpg, jpeg, tiff, tif, bmp, svg).</param>
    /// <param name="options">Conversion options including page index and DPI.</param>
    /// <exception cref="ArgumentException">Thrown when the format is not supported.</exception>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when pageIndex is out of range.</exception>
    public static void ConvertWordToImages(Document document, string outputPath, string outputFormat,
        ConversionOptions options)
    {
        var format = NormalizeExtension(outputFormat);
        var saveFormat = format switch
        {
            "png" => WordSaveFormat.Png,
            "jpg" or "jpeg" => WordSaveFormat.Jpeg,
            "tiff" or "tif" => WordSaveFormat.Tiff,
            "bmp" => WordSaveFormat.Bmp,
            "svg" => WordSaveFormat.Svg,
            _ => throw new ArgumentException($"Unsupported image format for Word: {format}")
        };

        if (options.PageIndex.HasValue)
        {
            if (options.PageIndex.Value < 1 || options.PageIndex.Value > document.PageCount)
                throw new ArgumentOutOfRangeException(nameof(options),
                    $"Page index must be between 1 and {document.PageCount}");

            var imageOptions = CreateWordImageSaveOptions(saveFormat, options);
            imageOptions.PageSet = new PageSet(options.PageIndex.Value - 1);
            document.Save(outputPath, imageOptions);
        }
        else
        {
            var dir = Path.GetDirectoryName(outputPath) ?? ".";
            var baseName = Path.GetFileNameWithoutExtension(outputPath);
            var ext = Path.GetExtension(outputPath);

            for (var i = 0; i < document.PageCount; i++)
            {
                var imageOptions = CreateWordImageSaveOptions(saveFormat, options);
                imageOptions.PageSet = new PageSet(i);
                var pagePath = document.PageCount == 1
                    ? outputPath
                    : Path.Combine(dir, $"{baseName}_{i + 1}{ext}");
                document.Save(pagePath, imageOptions);
            }
        }
    }

    /// <summary>
    ///     Creates Word image save options with JPEG quality support.
    /// </summary>
    /// <param name="saveFormat">The image save format.</param>
    /// <param name="options">Conversion options.</param>
    /// <returns>Configured image save options.</returns>
    private static ImageSaveOptions CreateWordImageSaveOptions(WordSaveFormat saveFormat, ConversionOptions options)
    {
        var imageOptions = new ImageSaveOptions(saveFormat)
        {
            Resolution = options.Dpi
        };

        if (saveFormat == WordSaveFormat.Jpeg)
            imageOptions.JpegQuality = options.JpegQuality;

        return imageOptions;
    }

    /// <summary>
    ///     Converts Excel workbook sheets to images.
    /// </summary>
    /// <param name="workbook">The Excel workbook to convert.</param>
    /// <param name="outputPath">The output file path (sheet number will be appended for multi-sheet output).</param>
    /// <param name="outputFormat">The target image format (png, jpg, jpeg, tiff, tif, bmp, svg).</param>
    /// <param name="options">Conversion options including sheet index and DPI.</param>
    /// <exception cref="ArgumentException">Thrown when the format is not supported.</exception>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when sheetIndex is out of range.</exception>
    public static void ConvertExcelToImages(Workbook workbook, string outputPath, string outputFormat,
        ConversionOptions options)
    {
        var format = NormalizeExtension(outputFormat);
        var imageType = format switch
        {
            "png" => ImageType.Png,
            "jpg" or "jpeg" => ImageType.Jpeg,
            "tiff" or "tif" => ImageType.Tiff,
            "bmp" => ImageType.Bmp,
            "svg" => ImageType.Svg,
            _ => throw new ArgumentException($"Unsupported image format for Excel: {format}")
        };

        var imageOptions = new ImageOrPrintOptions
        {
            ImageType = imageType,
            HorizontalResolution = options.Dpi,
            VerticalResolution = options.Dpi,
            OnePagePerSheet = true
        };

        if (imageType == ImageType.Jpeg)
            imageOptions.Quality = options.JpegQuality;

        if (options.PageIndex.HasValue)
        {
            if (options.PageIndex.Value < 1 || options.PageIndex.Value > workbook.Worksheets.Count)
                throw new ArgumentOutOfRangeException(nameof(options),
                    $"Sheet index must be between 1 and {workbook.Worksheets.Count}");

            var sheet = workbook.Worksheets[options.PageIndex.Value - 1];
            var sr = new SheetRender(sheet, imageOptions);
            sr.ToImage(0, outputPath);
        }
        else
        {
            var dir = Path.GetDirectoryName(outputPath) ?? ".";
            var baseName = Path.GetFileNameWithoutExtension(outputPath);
            var ext = Path.GetExtension(outputPath);

            for (var i = 0; i < workbook.Worksheets.Count; i++)
            {
                var sheet = workbook.Worksheets[i];
                var sr = new SheetRender(sheet, imageOptions);
                var sheetPath = workbook.Worksheets.Count == 1
                    ? outputPath
                    : Path.Combine(dir, $"{baseName}_{i + 1}{ext}");
                sr.ToImage(0, sheetPath);
            }
        }
    }

    /// <summary>
    ///     Converts a special format file (HTML, EPUB, Markdown, SVG, XPS, LaTeX, MHT) to PDF.
    /// </summary>
    /// <param name="inputPath">The input file path.</param>
    /// <param name="outputPath">The output PDF file path.</param>
    /// <returns>The source format name for result reporting.</returns>
    /// <exception cref="ArgumentException">Thrown when the input format is not supported.</exception>
    public static string ConvertToPdfFromSpecialFormat(string inputPath, string outputPath)
    {
        var extension = NormalizeExtension(Path.GetExtension(inputPath));

        switch (extension)
        {
            case "html":
            case "htm":
                using (var pdfDoc = new Aspose.Pdf.Document(inputPath, new HtmlLoadOptions()))
                {
                    pdfDoc.Save(outputPath);
                }

                return "HTML";

            case "epub":
                using (var pdfDoc = new Aspose.Pdf.Document(inputPath, new EpubLoadOptions()))
                {
                    pdfDoc.Save(outputPath);
                }

                return "EPUB";

            case "md":
                using (var pdfDoc = new Aspose.Pdf.Document(inputPath, new MdLoadOptions()))
                {
                    pdfDoc.Save(outputPath);
                }

                return "Markdown";

            case "svg":
                using (var pdfDoc = new Aspose.Pdf.Document(inputPath, new SvgLoadOptions()))
                {
                    pdfDoc.Save(outputPath);
                }

                return "SVG";

            case "xps":
                using (var pdfDoc = new Aspose.Pdf.Document(inputPath, new XpsLoadOptions()))
                {
                    pdfDoc.Save(outputPath);
                }

                return "XPS";

            case "tex":
                using (var pdfDoc = new Aspose.Pdf.Document(inputPath, new TeXLoadOptions()))
                {
                    pdfDoc.Save(outputPath);
                }

                return "LaTeX";

            case "mht":
            case "mhtml":
                using (var pdfDoc = new Aspose.Pdf.Document(inputPath, new MhtLoadOptions()))
                {
                    pdfDoc.Save(outputPath);
                }

                return "MHT";

            default:
                throw new ArgumentException($"Unsupported format for PDF conversion: {extension}");
        }
    }

    #endregion

    #region Format Support

    /// <summary>
    ///     Gets the MIME type for the specified output format.
    /// </summary>
    /// <param name="outputFormat">The output format (e.g., "pdf", "html", "png").</param>
    /// <returns>The MIME type string, or "application/octet-stream" if not found.</returns>
    public static string GetMimeType(string outputFormat)
    {
        var format = NormalizeExtension(outputFormat);
        return MimeTypes.GetValueOrDefault(format, "application/octet-stream");
    }

    /// <summary>
    ///     Checks whether the specified output format is supported for the given document type.
    /// </summary>
    /// <param name="documentType">The document type to check.</param>
    /// <param name="outputFormat">The output format to check.</param>
    /// <returns><c>true</c> if the format is supported; otherwise, <c>false</c>.</returns>
    public static bool IsFormatSupported(DocumentType documentType, string outputFormat)
    {
        var format = NormalizeExtension(outputFormat);
        return SupportedFormats.TryGetValue(documentType, out var formats) && formats.Contains(format);
    }

    /// <summary>
    ///     Gets all supported output formats for the specified document type.
    /// </summary>
    /// <param name="documentType">The document type.</param>
    /// <returns>An enumerable of supported format strings.</returns>
    public static IEnumerable<string> GetSupportedFormats(DocumentType documentType)
    {
        return SupportedFormats.TryGetValue(documentType, out var formats)
            ? formats
            : Enumerable.Empty<string>();
    }

    #endregion

    #region Private Methods

    /// <summary>
    ///     Normalizes the format/extension string by removing leading dots and converting to lowercase.
    /// </summary>
    /// <param name="format">The format string to normalize.</param>
    /// <returns>The normalized format string.</returns>
    private static string NormalizeExtension(string format)
    {
        return format.TrimStart('.').ToLowerInvariant();
    }

    /// <summary>
    ///     Internal method that performs the actual conversion to a stream.
    /// </summary>
    /// <param name="document">The source document.</param>
    /// <param name="documentType">The document type.</param>
    /// <param name="format">The normalized output format.</param>
    /// <param name="outputStream">The stream to write the converted document to.</param>
    /// <param name="progress">Optional progress reporter.</param>
    /// <param name="options">Optional conversion options.</param>
    /// <exception cref="ArgumentException">Thrown when the document type is not supported.</exception>
    private static void ConvertToStreamInternal(object document, DocumentType documentType, string format,
        Stream outputStream, IProgress<ProgressNotificationValue>? progress, ConversionOptions? options)
    {
        options ??= new ConversionOptions();

        switch (documentType)
        {
            case DocumentType.Word:
                ConvertWordToStream((Document)document, format, outputStream, progress, options);
                break;

            case DocumentType.Excel:
                ConvertExcelToStream((Workbook)document, format, outputStream, progress, options);
                break;

            case DocumentType.PowerPoint:
                ConvertPowerPointToStream((Presentation)document, format, outputStream, progress, options);
                break;

            case DocumentType.Pdf:
                ConvertPdfToStream((Aspose.Pdf.Document)document, format, outputStream, options);
                break;

            default:
                throw new ArgumentException($"Unsupported document type: {documentType}");
        }
    }

    /// <summary>
    ///     Converts a Word document to a stream.
    ///     For image formats, renders only the first page.
    /// </summary>
    /// <param name="document">The Word document to convert.</param>
    /// <param name="format">The normalized output format.</param>
    /// <param name="outputStream">The stream to write the converted document to.</param>
    /// <param name="progress">Optional progress reporter.</param>
    /// <param name="options">Conversion options.</param>
    /// <exception cref="ArgumentException">Thrown when the output format is not supported.</exception>
    private static void ConvertWordToStream(Document document, string format, Stream outputStream,
        IProgress<ProgressNotificationValue>? progress, ConversionOptions options)
    {
        document.UpdatePageLayout();

        if (IsWordImageFormat(format))
        {
            var saveFormat = format switch
            {
                "png" => WordSaveFormat.Png,
                "jpg" or "jpeg" => WordSaveFormat.Jpeg,
                "tiff" or "tif" => WordSaveFormat.Tiff,
                "bmp" => WordSaveFormat.Bmp,
                "svg" => WordSaveFormat.Svg,
                _ => throw new ArgumentException($"Unsupported image format for Word: {format}")
            };
            var imageOptions = new ImageSaveOptions(saveFormat) { PageSet = new PageSet(0) };
            if (saveFormat == WordSaveFormat.Jpeg)
                imageOptions.JpegQuality = options.JpegQuality;
            document.Save(outputStream, imageOptions);
            return;
        }

        if (format == "pdf")
        {
            var saveOptions = new PdfSaveOptions
            {
                ProgressCallback = new WordsProgressAdapter(progress)
            };
            ApplyWordPdfCompliance(saveOptions, options.PdfCompliance);
            document.Save(outputStream, saveOptions);
        }
        else if (format is "html" or "htm")
        {
            var saveOptions = new WordHtmlSaveOptions
            {
                ExportImagesAsBase64 = options.HtmlEmbedImages || options.HtmlSingleFile,
                ExportFontsAsBase64 = options.HtmlSingleFile
            };
            document.Save(outputStream, saveOptions);
        }
        else
        {
            var saveFormat = GetWordSaveFormat(format);
            document.Save(outputStream, saveFormat);
        }
    }

    /// <summary>
    ///     Converts an Excel workbook to a stream.
    ///     For image formats, renders only the first sheet.
    /// </summary>
    /// <param name="workbook">The Excel workbook to convert.</param>
    /// <param name="format">The normalized output format.</param>
    /// <param name="outputStream">The stream to write the converted document to.</param>
    /// <param name="progress">Optional progress reporter.</param>
    /// <param name="options">Conversion options.</param>
    /// <exception cref="ArgumentException">Thrown when the output format is not supported.</exception>
    private static void ConvertExcelToStream(Workbook workbook, string format, Stream outputStream,
        IProgress<ProgressNotificationValue>? progress, ConversionOptions options)
    {
        workbook.CalculateFormula();

        if (IsExcelImageFormat(format))
        {
            var imageType = format switch
            {
                "png" => ImageType.Png,
                "jpg" or "jpeg" => ImageType.Jpeg,
                "tiff" or "tif" => ImageType.Tiff,
                "bmp" => ImageType.Bmp,
                "svg" => ImageType.Svg,
                _ => throw new ArgumentException($"Unsupported image format for Excel: {format}")
            };

            var imageOptions = new ImageOrPrintOptions
            {
                ImageType = imageType,
                OnePagePerSheet = true
            };
            if (imageType == ImageType.Jpeg)
                imageOptions.Quality = options.JpegQuality;

            var sr = new SheetRender(workbook.Worksheets[0], imageOptions);
            sr.ToImage(0, outputStream);
            return;
        }

        if (format == "pdf")
        {
            var saveOptions = new Aspose.Cells.PdfSaveOptions
            {
                PageSavingCallback = new CellsProgressAdapter(progress)
            };
            ApplyExcelPdfCompliance(saveOptions, options.PdfCompliance);
            workbook.Save(outputStream, saveOptions);
        }
        else if (format is "html" or "htm")
        {
            var saveOptions = new HtmlSaveOptions
            {
                ExportImagesAsBase64 = options.HtmlEmbedImages || options.HtmlSingleFile,
                SaveAsSingleFile = options.HtmlSingleFile,
                ShowAllSheets = options.HtmlSingleFile
            };
            workbook.Save(outputStream, saveOptions);
        }
        else if (format == "csv")
        {
            var saveOptions = new TxtSaveOptions
            {
                Separator = string.IsNullOrEmpty(options.CsvSeparator) ? ',' : options.CsvSeparator[0]
            };
            workbook.Save(outputStream, saveOptions);
        }
        else
        {
            var saveFormat = GetExcelSaveFormat(format);
            workbook.Save(outputStream, saveFormat);
        }
    }

    /// <summary>
    ///     Converts a PowerPoint presentation to a stream.
    /// </summary>
    /// <param name="presentation">The PowerPoint presentation to convert.</param>
    /// <param name="format">The normalized output format.</param>
    /// <param name="outputStream">The stream to write the converted document to.</param>
    /// <param name="progress">Optional progress reporter.</param>
    /// <param name="options">Conversion options.</param>
    /// <exception cref="ArgumentException">Thrown when the output format is not supported.</exception>
    /// <exception cref="InvalidOperationException">Thrown when the presentation has no slides (for image output).</exception>
    private static void ConvertPowerPointToStream(Presentation presentation, string format, Stream outputStream,
        IProgress<ProgressNotificationValue>? progress, ConversionOptions options)
    {
        if (format is "png" or "jpg" or "jpeg")
        {
            ConvertPresentationToImageStream(presentation, format, outputStream, options);
            return;
        }

        if (format == "pdf")
        {
            var saveOptions = new PdfOptions
            {
                ProgressCallback = new SlidesProgressAdapter(progress),
                JpegQuality = (byte)options.JpegQuality
            };
            ApplySlidesPdfCompliance(saveOptions, options.PdfCompliance);
            presentation.Save(outputStream, Aspose.Slides.Export.SaveFormat.Pdf, saveOptions);
        }
        else
        {
            var saveFormat = GetPresentationSaveFormat(format);
            presentation.Save(outputStream, saveFormat);
        }
    }

    /// <summary>
    ///     Converts a PowerPoint presentation to an image stream.
    ///     For multi-slide presentations, renders the first slide.
    /// </summary>
    /// <param name="presentation">The PowerPoint presentation to convert.</param>
    /// <param name="format">The normalized image format (png, jpg, jpeg).</param>
    /// <param name="outputStream">The stream to write the image to.</param>
    /// <param name="options">Conversion options.</param>
    /// <exception cref="InvalidOperationException">Thrown when the presentation has no slides.</exception>
    /// <exception cref="ArgumentException">Thrown when the image format is not supported.</exception>
    private static void ConvertPresentationToImageStream(Presentation presentation, string format, Stream outputStream,
        ConversionOptions options)
    {
        if (presentation.Slides.Count == 0)
            throw new InvalidOperationException("Presentation has no slides to convert.");

        var slide = presentation.Slides[0];

        if (format is "jpg" or "jpeg")
        {
            var encoder = ImageCodecInfo.GetImageEncoders().First(c => c.FormatID == ImageFormat.Jpeg.Guid);
            var encoderParams = new EncoderParameters(1);
            encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, options.JpegQuality);

            using var bitmap = slide.GetThumbnail(1f, 1f);
            bitmap.Save(outputStream, encoder, encoderParams);
        }
        else
        {
            var imageFormat = format switch
            {
                "png" => ImageFormat.Png,
                _ => throw new ArgumentException($"Unsupported image format: {format}")
            };

            using var bitmap = slide.GetThumbnail(1f, 1f);
            bitmap.Save(outputStream, imageFormat);
        }
    }

    /// <summary>
    ///     Converts a PDF document to a stream.
    ///     For image formats, renders only the first page.
    /// </summary>
    /// <param name="pdfDocument">The PDF document to convert.</param>
    /// <param name="format">The normalized output format.</param>
    /// <param name="outputStream">The stream to write the converted document to.</param>
    /// <param name="options">Conversion options.</param>
    /// <exception cref="ArgumentException">Thrown when the output format is not supported.</exception>
    /// <exception cref="InvalidOperationException">Thrown when the PDF has no pages (for image output).</exception>
    private static void ConvertPdfToStream(Aspose.Pdf.Document pdfDocument, string format, Stream outputStream,
        ConversionOptions options)
    {
        if (IsImageFormat(format))
        {
            ConvertPdfFirstPageToImageStream(pdfDocument, format, outputStream, options);
        }
        else if (format is "html" or "htm")
        {
            var htmlOptions = new Aspose.Pdf.HtmlSaveOptions
            {
                PartsEmbeddingMode = Aspose.Pdf.HtmlSaveOptions.PartsEmbeddingModes.EmbedAllIntoHtml,
                RasterImagesSavingMode =
                    Aspose.Pdf.HtmlSaveOptions.RasterImagesSavingModes.AsEmbeddedPartsOfPngPageBackground,
                FontSavingMode = Aspose.Pdf.HtmlSaveOptions.FontSavingModes.SaveInAllFormats
            };
            pdfDocument.Save(outputStream, htmlOptions);
        }
        else if (format == "svg")
        {
            var svgOptions = new SvgSaveOptions
            {
                CompressOutputToZipArchive = false
            };
            pdfDocument.Save(outputStream, svgOptions);
        }
        else
        {
            var saveFormat = GetPdfSaveFormat(format);
            pdfDocument.Save(outputStream, saveFormat);
        }
    }

    /// <summary>
    ///     Converts the first page of a PDF document to an image stream.
    /// </summary>
    /// <param name="pdfDocument">The PDF document to convert.</param>
    /// <param name="format">The normalized image format (png, jpg, jpeg, tiff, tif).</param>
    /// <param name="outputStream">The stream to write the image to.</param>
    /// <param name="options">Conversion options.</param>
    /// <exception cref="InvalidOperationException">Thrown when the PDF has no pages.</exception>
    /// <exception cref="ArgumentException">Thrown when the image format is not supported.</exception>
    private static void ConvertPdfFirstPageToImageStream(Aspose.Pdf.Document pdfDocument, string format,
        Stream outputStream, ConversionOptions options)
    {
        if (pdfDocument.Pages.Count == 0)
            throw new InvalidOperationException("PDF document has no pages to convert.");

        var resolution = new Resolution(options.Dpi);

        if (format is "tiff" or "tif")
        {
            var tiffDevice = new TiffDevice(resolution);
            tiffDevice.Process(pdfDocument, outputStream);
            return;
        }

        PageDevice device = format switch
        {
            "png" => new PngDevice(resolution),
            "jpg" or "jpeg" => new JpegDevice(resolution, options.JpegQuality),
            _ => throw new ArgumentException($"Unsupported image format: {format}")
        };

        device.Process(pdfDocument.Pages[1], outputStream);
    }

    #endregion
}
