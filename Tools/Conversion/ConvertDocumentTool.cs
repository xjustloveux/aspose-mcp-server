using System.ComponentModel;
using Aspose.Cells;
using Aspose.Pdf;
using Aspose.Pdf.Devices;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Progress;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Conversion;
using ModelContextProtocol;
using ModelContextProtocol.Server;
using PdfSaveOptions = Aspose.Words.Saving.PdfSaveOptions;
using SaveFormat = Aspose.Words.SaveFormat;

namespace AsposeMcpServer.Tools.Conversion;

/// <summary>
///     Tool for converting documents between various formats with automatic source type detection.
///     Supports Word, Excel, PowerPoint, and PDF documents.
///     PDF output includes document formats (DOCX, HTML, XLSX, PPTX, EPUB, SVG, XPS, XML)
///     and image formats (PNG, JPEG, TIFF) with per-page rendering.
/// </summary>
[McpServerToolType]
public class ConvertDocumentTool
{
    /// <summary>
    ///     The session identity accessor for session isolation.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     The document session manager for managing in-memory document sessions.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ConvertDocumentTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public ConvertDocumentTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Converts a document between various formats with automatic source type detection.
    /// </summary>
    /// <param name="inputPath">Input file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID to convert document from session.</param>
    /// <param name="outputPath">Output file path (required, format determined by extension).</param>
    /// <param name="progress">Optional progress reporter for long-running operations.</param>
    /// <returns>A ConversionResult indicating the conversion result with source and output information.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when outputPath is not provided, neither inputPath nor sessionId is provided, or the input format is
    ///     unsupported.
    /// </exception>
    /// <exception cref="InvalidOperationException">Thrown when session management is not enabled but sessionId is provided.</exception>
    /// <exception cref="KeyNotFoundException">Thrown when the specified session is not found or access is denied.</exception>
    [McpServerTool(
        Name = "convert_document",
        Title = "Convert Document Between Formats",
        Destructive = false,
        Idempotent = true,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [OutputSchema(typeof(ConversionResult))]
    [Description(@"Convert documents between various formats (auto-detect source type).
Supports Word, Excel, PowerPoint, and PDF as input.

Usage examples:
- Convert Word to HTML: convert_document(inputPath='doc.docx', outputPath='doc.html')
- Convert Excel to CSV: convert_document(inputPath='book.xlsx', outputPath='book.csv')
- Convert PowerPoint to PDF: convert_document(inputPath='presentation.pptx', outputPath='presentation.pdf')
- Convert PDF to Word: convert_document(inputPath='document.pdf', outputPath='document.docx')
- Convert PDF to Excel: convert_document(inputPath='data.pdf', outputPath='data.xlsx')
- Convert PDF to PowerPoint: convert_document(inputPath='slides.pdf', outputPath='slides.pptx')
- Convert PDF to images: convert_document(inputPath='doc.pdf', outputPath='page.png')
- Convert PDF to EPUB: convert_document(inputPath='doc.pdf', outputPath='doc.epub')
- Convert PDF to SVG: convert_document(inputPath='doc.pdf', outputPath='doc.svg')
- Convert from session: convert_document(sessionId='sess_xxx', outputPath='doc.pdf')")]
    public ConversionResult Execute(
        [Description("Input file path (required if no sessionId)")]
        string? inputPath = null,
        [Description("Session ID to convert document from session")]
        string? sessionId = null,
        [Description("Output file path (required, format determined by extension)")]
        string? outputPath = null,
        IProgress<ProgressNotificationValue>? progress = null)
    {
        if (string.IsNullOrEmpty(outputPath))
            throw new ArgumentException("outputPath is required");

        SecurityHelper.ValidateFilePath(outputPath, nameof(outputPath), true);

        if (string.IsNullOrEmpty(inputPath) && string.IsNullOrEmpty(sessionId))
            throw new ArgumentException("Either inputPath or sessionId must be provided");

        var outputExtension = Path.GetExtension(outputPath).ToLower();

        if (!string.IsNullOrEmpty(sessionId))
            return ConvertFromSession(sessionId, outputPath, outputExtension, $"session:{sessionId}", progress);

        SecurityHelper.ValidateFilePath(inputPath!, nameof(inputPath), true);
        return ConvertFromFile(inputPath!, outputPath, outputExtension, progress);
    }

    /// <summary>
    ///     Converts a document from a session to the specified output format.
    /// </summary>
    /// <param name="sessionId">The session ID containing the document.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="outputExtension">The target format extension (e.g., ".pdf", ".html").</param>
    /// <param name="sourcePath">The source path for the result.</param>
    /// <param name="progress">Optional progress reporter for long-running operations.</param>
    /// <returns>A ConversionResult indicating the conversion result.</returns>
    /// <exception cref="InvalidOperationException">Thrown when session management is not enabled.</exception>
    /// <exception cref="KeyNotFoundException">Thrown when the session is not found or access is denied.</exception>
    /// <exception cref="ArgumentException">Thrown when the document type is unsupported.</exception>
    private ConversionResult ConvertFromSession(string sessionId, string outputPath, string outputExtension,
        string sourcePath, IProgress<ProgressNotificationValue>? progress)
    {
        if (_sessionManager == null)
            throw new InvalidOperationException("Session management is not enabled");

        var identity = _identityAccessor?.GetCurrentIdentity() ?? SessionIdentity.GetAnonymous();
        var session = _sessionManager.TryGetSession(sessionId, identity)
                      ?? throw new KeyNotFoundException($"Session '{sessionId}' not found or access denied");

        string sourceType;
        switch (session.Type)
        {
            case DocumentType.Word:
                var wordDoc = _sessionManager.GetDocument<Aspose.Words.Document>(sessionId, identity);
                var wordFormat = GetWordSaveFormat(outputExtension);
                if (outputExtension == ".pdf")
                {
                    var wordSaveOptions = new PdfSaveOptions
                    {
                        ProgressCallback = new WordsProgressAdapter(progress)
                    };
                    wordDoc.Save(outputPath, wordSaveOptions);
                }
                else
                {
                    wordDoc.Save(outputPath, wordFormat);
                }

                sourceType = "Word";
                break;

            case DocumentType.Excel:
                var workbook = _sessionManager.GetDocument<Workbook>(sessionId, identity);
                var excelFormat = GetExcelSaveFormat(outputExtension);
                if (outputExtension == ".pdf")
                {
                    var cellsSaveOptions = new Aspose.Cells.PdfSaveOptions
                    {
                        PageSavingCallback = new CellsProgressAdapter(progress)
                    };
                    workbook.Save(outputPath, cellsSaveOptions);
                }
                else
                {
                    workbook.Save(outputPath, excelFormat);
                }

                sourceType = "Excel";
                break;

            case DocumentType.PowerPoint:
                var presentation = _sessionManager.GetDocument<Presentation>(sessionId, identity);
                var pptFormat = GetPresentationSaveFormat(outputExtension);
                if (outputExtension == ".pdf")
                {
                    var slidesSaveOptions = new PdfOptions
                    {
                        ProgressCallback = new SlidesProgressAdapter(progress)
                    };
                    presentation.Save(outputPath, Aspose.Slides.Export.SaveFormat.Pdf, slidesSaveOptions);
                }
                else
                {
                    presentation.Save(outputPath, pptFormat);
                }

                sourceType = "PowerPoint";
                break;

            case DocumentType.Pdf:
                var pdfDoc = _sessionManager.GetDocument<Document>(sessionId, identity);
                if (IsImageExtension(outputExtension))
                {
                    ConvertPdfToImages(pdfDoc, outputPath, outputExtension);
                }
                else
                {
                    var pdfFormat = GetPdfSaveFormat(outputExtension);
                    pdfDoc.Save(outputPath, pdfFormat);
                }

                sourceType = "PDF";
                break;

            default:
                throw new ArgumentException($"Unsupported document type: {session.Type}");
        }

        return new ConversionResult
        {
            SourcePath = sourcePath,
            OutputPath = outputPath,
            SourceFormat = sourceType,
            TargetFormat = outputExtension.TrimStart('.').ToUpperInvariant(),
            FileSize = File.Exists(outputPath) ? new FileInfo(outputPath).Length : null,
            Message = $"Document from session {sessionId} ({sourceType}) converted to {outputExtension} format"
        };
    }

    /// <summary>
    ///     Converts a document file to the specified output format.
    /// </summary>
    /// <param name="inputPath">The source document file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="outputExtension">The target format extension (e.g., ".pdf", ".html").</param>
    /// <param name="progress">Optional progress reporter for long-running operations.</param>
    /// <returns>A ConversionResult indicating the conversion result.</returns>
    /// <exception cref="ArgumentException">Thrown when the input format is unsupported.</exception>
    private static ConversionResult ConvertFromFile(string inputPath, string outputPath, string outputExtension,
        IProgress<ProgressNotificationValue>? progress)
    {
        var inputExtension = Path.GetExtension(inputPath).ToLower();
        string sourceFormat;

        if (IsWordDocument(inputExtension))
        {
            var doc = new Aspose.Words.Document(inputPath);
            if (outputExtension == ".pdf")
            {
                var wordSaveOptions = new PdfSaveOptions
                {
                    ProgressCallback = new WordsProgressAdapter(progress)
                };
                doc.Save(outputPath, wordSaveOptions);
            }
            else
            {
                var saveFormat = GetWordSaveFormat(outputExtension);
                doc.Save(outputPath, saveFormat);
            }

            sourceFormat = "Word";
        }
        else if (IsExcelDocument(inputExtension))
        {
            using var workbook = new Workbook(inputPath);
            if (outputExtension == ".pdf")
            {
                var cellsSaveOptions = new Aspose.Cells.PdfSaveOptions
                {
                    PageSavingCallback = new CellsProgressAdapter(progress)
                };
                workbook.Save(outputPath, cellsSaveOptions);
            }
            else
            {
                var saveFormat = GetExcelSaveFormat(outputExtension);
                workbook.Save(outputPath, saveFormat);
            }

            sourceFormat = "Excel";
        }
        else if (IsPresentationDocument(inputExtension))
        {
            using var presentation = new Presentation(inputPath);
            if (outputExtension == ".pdf")
            {
                var slidesSaveOptions = new PdfOptions
                {
                    ProgressCallback = new SlidesProgressAdapter(progress)
                };
                presentation.Save(outputPath, Aspose.Slides.Export.SaveFormat.Pdf, slidesSaveOptions);
            }
            else
            {
                var saveFormat = GetPresentationSaveFormat(outputExtension);
                presentation.Save(outputPath, saveFormat);
            }

            sourceFormat = "PowerPoint";
        }
        else if (IsPdfDocument(inputExtension))
        {
            if (IsImageExtension(outputExtension))
            {
                ConvertPdfToImages(inputPath, outputPath, outputExtension);
            }
            else
            {
                using var pdfDoc = new Document(inputPath);
                var pdfFormat = GetPdfSaveFormat(outputExtension);
                pdfDoc.Save(outputPath, pdfFormat);
            }

            sourceFormat = "PDF";
        }
        else
        {
            throw new ArgumentException($"Unsupported input format: {inputExtension}");
        }

        return new ConversionResult
        {
            SourcePath = inputPath,
            OutputPath = outputPath,
            SourceFormat = sourceFormat,
            TargetFormat = outputExtension.TrimStart('.').ToUpperInvariant(),
            FileSize = File.Exists(outputPath) ? new FileInfo(outputPath).Length : null,
            Message = $"Document converted from {inputExtension} to {outputExtension} format"
        };
    }

    /// <summary>
    ///     Determines whether the specified extension represents a Word document.
    /// </summary>
    /// <param name="extension">The file extension to check.</param>
    /// <returns><c>true</c> if the extension is a Word document format; otherwise, <c>false</c>.</returns>
    private static bool IsWordDocument(string extension)
    {
        return extension is ".doc" or ".docx" or ".rtf" or ".odt" or ".txt";
    }

    /// <summary>
    ///     Determines whether the specified extension represents an Excel document.
    /// </summary>
    /// <param name="extension">The file extension to check.</param>
    /// <returns><c>true</c> if the extension is an Excel document format; otherwise, <c>false</c>.</returns>
    private static bool IsExcelDocument(string extension)
    {
        return extension is ".xls" or ".xlsx" or ".csv" or ".ods";
    }

    /// <summary>
    ///     Determines whether the specified extension represents a PowerPoint presentation.
    /// </summary>
    /// <param name="extension">The file extension to check.</param>
    /// <returns><c>true</c> if the extension is a PowerPoint format; otherwise, <c>false</c>.</returns>
    private static bool IsPresentationDocument(string extension)
    {
        return extension is ".ppt" or ".pptx" or ".odp";
    }

    /// <summary>
    ///     Determines whether the specified extension represents a PDF document.
    /// </summary>
    /// <param name="extension">The file extension to check.</param>
    /// <returns><c>true</c> if the extension is a PDF document format; otherwise, <c>false</c>.</returns>
    private static bool IsPdfDocument(string extension)
    {
        return extension is ".pdf";
    }

    /// <summary>
    ///     Determines whether the specified extension represents an image format.
    /// </summary>
    /// <param name="extension">The file extension to check.</param>
    /// <returns><c>true</c> if the extension is an image format; otherwise, <c>false</c>.</returns>
    private static bool IsImageExtension(string extension)
    {
        return extension is ".png" or ".jpg" or ".jpeg" or ".tiff" or ".tif";
    }

    /// <summary>
    ///     Converts a PDF file to images, one per page.
    /// </summary>
    /// <param name="inputPath">The input PDF file path.</param>
    /// <param name="outputPath">The output image file path (page number will be appended for multi-page).</param>
    /// <param name="extension">The target image extension.</param>
    private static void ConvertPdfToImages(string inputPath, string outputPath, string extension)
    {
        using var pdfDoc = new Document(inputPath);
        ConvertPdfToImages(pdfDoc, outputPath, extension);
    }

    /// <summary>
    ///     Converts a PDF document to images, one per page.
    ///     TIFF format uses <see cref="Aspose.Pdf.Devices.TiffDevice" /> which produces a single multi-page file.
    ///     PNG and JPEG formats use per-page devices, producing one file per page.
    /// </summary>
    /// <param name="pdfDoc">The PDF document to convert.</param>
    /// <param name="outputPath">The output image file path (page number will be appended for multi-page PNG/JPEG).</param>
    /// <param name="extension">The target image extension.</param>
    private static void ConvertPdfToImages(Document pdfDoc, string outputPath, string extension)
    {
        var resolution = new Resolution(150);

        if (extension is ".tiff" or ".tif")
        {
            var tiffDevice = new TiffDevice(resolution);
            using var stream = new FileStream(outputPath, FileMode.Create);
            tiffDevice.Process(pdfDoc, stream);
            return;
        }

        var dir = Path.GetDirectoryName(outputPath) ?? ".";
        var nameWithoutExt = Path.GetFileNameWithoutExtension(outputPath);
        var ext = Path.GetExtension(outputPath);

        for (var i = 1; i <= pdfDoc.Pages.Count; i++)
        {
            var pagePath = pdfDoc.Pages.Count == 1
                ? outputPath
                : Path.Combine(dir, $"{nameWithoutExt}_{i}{ext}");

            using var stream = new FileStream(pagePath, FileMode.Create);
            PageDevice device = extension switch
            {
                ".png" => new PngDevice(resolution),
                ".jpg" or ".jpeg" => new JpegDevice(resolution),
                _ => throw new ArgumentException($"Unsupported image format: {extension}")
            };

            device.Process(pdfDoc.Pages[i], stream);
        }
    }

    /// <summary>
    ///     Gets the Word save format for the specified extension.
    /// </summary>
    /// <param name="extension">The target file extension.</param>
    /// <returns>The corresponding <see cref="SaveFormat" /> value.</returns>
    /// <exception cref="ArgumentException">Thrown when the extension is not supported for Word output.</exception>
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
            _ => throw new ArgumentException($"Unsupported output format for Word: {extension}")
        };
    }

    /// <summary>
    ///     Gets the Excel save format for the specified extension.
    /// </summary>
    /// <param name="extension">The target file extension.</param>
    /// <returns>The corresponding <see cref="Aspose.Cells.SaveFormat" /> value.</returns>
    /// <exception cref="ArgumentException">Thrown when the extension is not supported for Excel output.</exception>
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
            _ => throw new ArgumentException($"Unsupported output format for Excel: {extension}")
        };
    }

    /// <summary>
    ///     Gets the PowerPoint save format for the specified extension.
    /// </summary>
    /// <param name="extension">The target file extension.</param>
    /// <returns>The corresponding <see cref="Aspose.Slides.Export.SaveFormat" /> value.</returns>
    /// <exception cref="ArgumentException">Thrown when the extension is not supported for PowerPoint output.</exception>
    private static Aspose.Slides.Export.SaveFormat GetPresentationSaveFormat(string extension)
    {
        return extension switch
        {
            ".pdf" => Aspose.Slides.Export.SaveFormat.Pdf,
            ".pptx" => Aspose.Slides.Export.SaveFormat.Pptx,
            ".ppt" => Aspose.Slides.Export.SaveFormat.Ppt,
            ".html" => Aspose.Slides.Export.SaveFormat.Html,
            ".odp" => Aspose.Slides.Export.SaveFormat.Odp,
            _ => throw new ArgumentException($"Unsupported output format for PowerPoint: {extension}")
        };
    }

    /// <summary>
    ///     Gets the PDF save format for the specified extension.
    /// </summary>
    /// <param name="extension">The target file extension.</param>
    /// <returns>The corresponding <see cref="Aspose.Pdf.SaveFormat" /> value.</returns>
    /// <exception cref="ArgumentException">Thrown when the extension is not supported for PDF output.</exception>
    private static Aspose.Pdf.SaveFormat GetPdfSaveFormat(string extension)
    {
        return extension switch
        {
            ".docx" => Aspose.Pdf.SaveFormat.DocX,
            ".doc" => Aspose.Pdf.SaveFormat.Doc,
            ".html" => Aspose.Pdf.SaveFormat.Html,
            ".xlsx" => Aspose.Pdf.SaveFormat.Excel,
            ".pptx" => Aspose.Pdf.SaveFormat.Pptx,
            ".txt" => Aspose.Pdf.SaveFormat.TeX,
            ".epub" => Aspose.Pdf.SaveFormat.Epub,
            ".svg" => Aspose.Pdf.SaveFormat.Svg,
            ".xps" => Aspose.Pdf.SaveFormat.Xps,
            ".xml" => Aspose.Pdf.SaveFormat.Xml,
            _ => throw new ArgumentException($"Unsupported output format for PDF: {extension}")
        };
    }
}
