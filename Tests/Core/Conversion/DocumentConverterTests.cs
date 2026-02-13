using Aspose.Cells;
using Aspose.Pdf.Text;
using Aspose.Slides;
using Aspose.Words;
using AsposeMcpServer.Core.Conversion;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Tests.Infrastructure;
using Document = Aspose.Words.Document;
using SaveFormat = Aspose.Words.SaveFormat;

namespace AsposeMcpServer.Tests.Core.Conversion;

/// <summary>
///     Unit tests for DocumentConverter class.
/// </summary>
public class DocumentConverterTests : TestBase
{
    #region ConvertToBytes Validation Tests

    [Fact]
    public void ConvertToBytes_NullDocument_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>(() =>
            DocumentConverter.ConvertToBytes(null!, DocumentType.Word, "pdf"));
    }

    #endregion

    #region GetMimeType Tests

    [Theory]
    [InlineData("pdf", "application/pdf")]
    [InlineData("PDF", "application/pdf")]
    [InlineData(".pdf", "application/pdf")]
    [InlineData("html", "text/html")]
    [InlineData("htm", "text/html")]
    [InlineData("png", "image/png")]
    [InlineData("jpg", "image/jpeg")]
    [InlineData("jpeg", "image/jpeg")]
    [InlineData("docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")]
    [InlineData("xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")]
    [InlineData("pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation")]
    [InlineData("csv", "text/csv")]
    [InlineData("epub", "application/epub+zip")]
    [InlineData("svg", "image/svg+xml")]
    public void GetMimeType_KnownFormats_ReturnsCorrectMimeType(string format, string expectedMimeType)
    {
        var mimeType = DocumentConverter.GetMimeType(format);

        Assert.Equal(expectedMimeType, mimeType);
    }

    [Theory]
    [InlineData("unknown")]
    [InlineData("xyz")]
    [InlineData("")]
    public void GetMimeType_UnknownFormats_ReturnsOctetStream(string format)
    {
        var mimeType = DocumentConverter.GetMimeType(format);

        Assert.Equal("application/octet-stream", mimeType);
    }

    #endregion

    #region IsFormatSupported Tests

    [Theory]
    [InlineData(DocumentType.Word, "pdf", true)]
    [InlineData(DocumentType.Word, "html", true)]
    [InlineData(DocumentType.Word, "docx", true)]
    [InlineData(DocumentType.Word, "doc", true)]
    [InlineData(DocumentType.Word, "rtf", true)]
    [InlineData(DocumentType.Word, "txt", true)]
    [InlineData(DocumentType.Word, "odt", true)]
    [InlineData(DocumentType.Word, "png", true)]
    [InlineData(DocumentType.Word, "jpg", true)]
    [InlineData(DocumentType.Word, "jpeg", true)]
    [InlineData(DocumentType.Word, "tiff", true)]
    [InlineData(DocumentType.Word, "tif", true)]
    [InlineData(DocumentType.Word, "bmp", true)]
    [InlineData(DocumentType.Word, "svg", true)]
    [InlineData(DocumentType.Word, "xlsx", false)]
    [InlineData(DocumentType.Word, "pptx", false)]
    public void IsFormatSupported_Word_ReturnsCorrectResult(DocumentType docType, string format, bool expected)
    {
        var result = DocumentConverter.IsFormatSupported(docType, format);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData(DocumentType.Excel, "pdf", true)]
    [InlineData(DocumentType.Excel, "html", true)]
    [InlineData(DocumentType.Excel, "xlsx", true)]
    [InlineData(DocumentType.Excel, "xls", true)]
    [InlineData(DocumentType.Excel, "csv", true)]
    [InlineData(DocumentType.Excel, "ods", true)]
    [InlineData(DocumentType.Excel, "png", true)]
    [InlineData(DocumentType.Excel, "jpg", true)]
    [InlineData(DocumentType.Excel, "jpeg", true)]
    [InlineData(DocumentType.Excel, "tiff", true)]
    [InlineData(DocumentType.Excel, "tif", true)]
    [InlineData(DocumentType.Excel, "bmp", true)]
    [InlineData(DocumentType.Excel, "svg", true)]
    [InlineData(DocumentType.Excel, "docx", false)]
    [InlineData(DocumentType.Excel, "pptx", false)]
    public void IsFormatSupported_Excel_ReturnsCorrectResult(DocumentType docType, string format, bool expected)
    {
        var result = DocumentConverter.IsFormatSupported(docType, format);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData(DocumentType.PowerPoint, "pdf", true)]
    [InlineData(DocumentType.PowerPoint, "html", true)]
    [InlineData(DocumentType.PowerPoint, "pptx", true)]
    [InlineData(DocumentType.PowerPoint, "ppt", true)]
    [InlineData(DocumentType.PowerPoint, "odp", true)]
    [InlineData(DocumentType.PowerPoint, "png", true)]
    [InlineData(DocumentType.PowerPoint, "jpg", true)]
    [InlineData(DocumentType.PowerPoint, "docx", false)]
    [InlineData(DocumentType.PowerPoint, "xlsx", false)]
    public void IsFormatSupported_PowerPoint_ReturnsCorrectResult(DocumentType docType, string format, bool expected)
    {
        var result = DocumentConverter.IsFormatSupported(docType, format);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData(DocumentType.Pdf, "docx", true)]
    [InlineData(DocumentType.Pdf, "doc", true)]
    [InlineData(DocumentType.Pdf, "html", true)]
    [InlineData(DocumentType.Pdf, "xlsx", true)]
    [InlineData(DocumentType.Pdf, "pptx", true)]
    [InlineData(DocumentType.Pdf, "png", true)]
    [InlineData(DocumentType.Pdf, "jpg", true)]
    [InlineData(DocumentType.Pdf, "jpeg", true)]
    [InlineData(DocumentType.Pdf, "tiff", true)]
    [InlineData(DocumentType.Pdf, "epub", true)]
    [InlineData(DocumentType.Pdf, "svg", true)]
    [InlineData(DocumentType.Pdf, "xps", true)]
    [InlineData(DocumentType.Pdf, "xml", true)]
    [InlineData(DocumentType.Pdf, "pdf", false)]
    [InlineData(DocumentType.Pdf, "odt", false)]
    public void IsFormatSupported_Pdf_ReturnsCorrectResult(DocumentType docType, string format, bool expected)
    {
        var result = DocumentConverter.IsFormatSupported(docType, format);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("PDF")]
    [InlineData("Pdf")]
    [InlineData(".pdf")]
    [InlineData(".PDF")]
    public void IsFormatSupported_CaseInsensitive_ReturnsTrue(string format)
    {
        var result = DocumentConverter.IsFormatSupported(DocumentType.Word, format);

        Assert.True(result);
    }

    [Theory]
    [InlineData("png", true)]
    [InlineData("jpg", true)]
    [InlineData("jpeg", true)]
    [InlineData("tiff", true)]
    [InlineData("tif", true)]
    [InlineData("bmp", true)]
    [InlineData("svg", true)]
    [InlineData("PNG", true)]
    [InlineData(".png", true)]
    [InlineData("pdf", false)]
    [InlineData("docx", false)]
    [InlineData("xlsx", false)]
    public void IsExcelImageFormat_ReturnsCorrectResult(string format, bool expected)
    {
        var result = DocumentConverter.IsExcelImageFormat(format);

        Assert.Equal(expected, result);
    }

    #endregion

    #region GetSupportedFormats Tests

    [Fact]
    public void GetSupportedFormats_Word_ReturnsExpectedFormats()
    {
        var formats = DocumentConverter.GetSupportedFormats(DocumentType.Word).ToList();

        Assert.Contains("pdf", formats);
        Assert.Contains("html", formats);
        Assert.Contains("docx", formats);
        Assert.Contains("doc", formats);
        Assert.Contains("rtf", formats);
        Assert.Contains("txt", formats);
        Assert.Contains("odt", formats);
        Assert.Contains("png", formats);
        Assert.Contains("jpg", formats);
        Assert.Contains("jpeg", formats);
        Assert.Contains("tiff", formats);
        Assert.Contains("tif", formats);
        Assert.Contains("bmp", formats);
        Assert.Contains("svg", formats);
        Assert.Equal(14, formats.Count);
    }

    [Fact]
    public void GetSupportedFormats_Excel_ReturnsExpectedFormats()
    {
        var formats = DocumentConverter.GetSupportedFormats(DocumentType.Excel).ToList();

        Assert.Contains("pdf", formats);
        Assert.Contains("html", formats);
        Assert.Contains("xlsx", formats);
        Assert.Contains("xls", formats);
        Assert.Contains("csv", formats);
        Assert.Contains("ods", formats);
        Assert.Contains("png", formats);
        Assert.Contains("jpg", formats);
        Assert.Contains("jpeg", formats);
        Assert.Contains("tiff", formats);
        Assert.Contains("tif", formats);
        Assert.Contains("bmp", formats);
        Assert.Contains("svg", formats);
        Assert.Equal(13, formats.Count);
    }

    [Fact]
    public void GetSupportedFormats_PowerPoint_ReturnsExpectedFormats()
    {
        var formats = DocumentConverter.GetSupportedFormats(DocumentType.PowerPoint).ToList();

        Assert.Contains("pdf", formats);
        Assert.Contains("html", formats);
        Assert.Contains("pptx", formats);
        Assert.Contains("ppt", formats);
        Assert.Contains("odp", formats);
        Assert.Contains("png", formats);
        Assert.Contains("jpg", formats);
        Assert.Equal(7, formats.Count);
    }

    [Fact]
    public void GetSupportedFormats_Pdf_ReturnsExpectedFormats()
    {
        var formats = DocumentConverter.GetSupportedFormats(DocumentType.Pdf).ToList();

        Assert.Contains("docx", formats);
        Assert.Contains("doc", formats);
        Assert.Contains("html", formats);
        Assert.Contains("xlsx", formats);
        Assert.Contains("pptx", formats);
        Assert.Contains("png", formats);
        Assert.Contains("jpg", formats);
        Assert.Contains("jpeg", formats);
        Assert.Contains("tiff", formats);
        Assert.Contains("tif", formats);
        Assert.Contains("epub", formats);
        Assert.Contains("svg", formats);
        Assert.Contains("xps", formats);
        Assert.Contains("xml", formats);
    }

    #endregion

    #region ConvertToStream Validation Tests

    [Fact]
    public void ConvertToStream_NullDocument_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>(() =>
            DocumentConverter.ConvertToStream(null!, DocumentType.Word, "pdf"));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void ConvertToStream_InvalidFormat_ThrowsArgumentException(string? format)
    {
        var doc = new Document();

        Assert.Throws<ArgumentException>(() =>
            DocumentConverter.ConvertToStream(doc, DocumentType.Word, format!));
    }

    [Fact]
    public void ConvertToStream_UnsupportedFormat_ThrowsArgumentException()
    {
        var doc = new Document();

        var ex = Assert.Throws<ArgumentException>(() =>
            DocumentConverter.ConvertToStream(doc, DocumentType.Word, "xlsx"));

        Assert.Contains("not supported", ex.Message);
    }

    #endregion

    #region Document Type Detection Tests

    [Theory]
    [InlineData("doc", true)]
    [InlineData("docx", true)]
    [InlineData("rtf", true)]
    [InlineData("odt", true)]
    [InlineData("txt", true)]
    [InlineData(".docx", true)]
    [InlineData("DOCX", true)]
    [InlineData("xlsx", false)]
    [InlineData("pdf", false)]
    [InlineData("pptx", false)]
    public void IsWordDocument_ReturnsCorrectResult(string extension, bool expected)
    {
        var result = DocumentConverter.IsWordDocument(extension);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("xls", true)]
    [InlineData("xlsx", true)]
    [InlineData("csv", true)]
    [InlineData("ods", true)]
    [InlineData(".xlsx", true)]
    [InlineData("XLSX", true)]
    [InlineData("docx", false)]
    [InlineData("pdf", false)]
    [InlineData("pptx", false)]
    public void IsExcelDocument_ReturnsCorrectResult(string extension, bool expected)
    {
        var result = DocumentConverter.IsExcelDocument(extension);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("ppt", true)]
    [InlineData("pptx", true)]
    [InlineData("odp", true)]
    [InlineData(".pptx", true)]
    [InlineData("PPTX", true)]
    [InlineData("docx", false)]
    [InlineData("pdf", false)]
    [InlineData("xlsx", false)]
    public void IsPowerPointDocument_ReturnsCorrectResult(string extension, bool expected)
    {
        var result = DocumentConverter.IsPowerPointDocument(extension);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("pdf", true)]
    [InlineData(".pdf", true)]
    [InlineData("PDF", true)]
    [InlineData("docx", false)]
    [InlineData("xlsx", false)]
    [InlineData("pptx", false)]
    public void IsPdfDocument_ReturnsCorrectResult(string extension, bool expected)
    {
        var result = DocumentConverter.IsPdfDocument(extension);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("png", true)]
    [InlineData("jpg", true)]
    [InlineData("jpeg", true)]
    [InlineData("tiff", true)]
    [InlineData("tif", true)]
    [InlineData(".png", true)]
    [InlineData("PNG", true)]
    [InlineData("pdf", false)]
    [InlineData("docx", false)]
    [InlineData("bmp", false)]
    [InlineData("svg", false)]
    public void IsImageFormat_ReturnsCorrectResult(string extension, bool expected)
    {
        var result = DocumentConverter.IsImageFormat(extension);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("html", true)]
    [InlineData("htm", true)]
    [InlineData("epub", true)]
    [InlineData("md", true)]
    [InlineData("svg", true)]
    [InlineData("xps", true)]
    [InlineData("tex", true)]
    [InlineData("mht", true)]
    [InlineData("mhtml", true)]
    [InlineData(".html", true)]
    [InlineData("HTML", true)]
    [InlineData("pdf", false)]
    [InlineData("docx", false)]
    [InlineData("xlsx", false)]
    public void IsPdfConvertibleFormat_ReturnsCorrectResult(string extension, bool expected)
    {
        var result = DocumentConverter.IsPdfConvertibleFormat(extension);

        Assert.Equal(expected, result);
    }

    #endregion

    #region GetDocumentType Tests

    [Theory]
    [InlineData("doc", DocumentType.Word)]
    [InlineData("docx", DocumentType.Word)]
    [InlineData("rtf", DocumentType.Word)]
    [InlineData("odt", DocumentType.Word)]
    [InlineData("txt", DocumentType.Word)]
    public void GetDocumentType_WordExtensions_ReturnsWord(string extension, DocumentType expected)
    {
        var result = DocumentConverter.GetDocumentType(extension);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("xls", DocumentType.Excel)]
    [InlineData("xlsx", DocumentType.Excel)]
    [InlineData("csv", DocumentType.Excel)]
    [InlineData("ods", DocumentType.Excel)]
    public void GetDocumentType_ExcelExtensions_ReturnsExcel(string extension, DocumentType expected)
    {
        var result = DocumentConverter.GetDocumentType(extension);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("ppt", DocumentType.PowerPoint)]
    [InlineData("pptx", DocumentType.PowerPoint)]
    [InlineData("odp", DocumentType.PowerPoint)]
    public void GetDocumentType_PowerPointExtensions_ReturnsPowerPoint(string extension, DocumentType expected)
    {
        var result = DocumentConverter.GetDocumentType(extension);

        Assert.Equal(expected, result);
    }

    [Fact]
    public void GetDocumentType_PdfExtension_ReturnsPdf()
    {
        var result = DocumentConverter.GetDocumentType("pdf");

        Assert.Equal(DocumentType.Pdf, result);
    }

    [Theory]
    [InlineData("html")]
    [InlineData("png")]
    [InlineData("jpg")]
    [InlineData("unknown")]
    public void GetDocumentType_UnsupportedExtensions_ReturnsNull(string extension)
    {
        var result = DocumentConverter.GetDocumentType(extension);

        Assert.Null(result);
    }

    #endregion

    #region GetWordSaveFormat Tests

    [Theory]
    [InlineData("pdf", SaveFormat.Pdf)]
    [InlineData("docx", SaveFormat.Docx)]
    [InlineData("doc", SaveFormat.Doc)]
    [InlineData("rtf", SaveFormat.Rtf)]
    [InlineData("html", SaveFormat.Html)]
    [InlineData("txt", SaveFormat.Text)]
    [InlineData("odt", SaveFormat.Odt)]
    [InlineData(".pdf", SaveFormat.Pdf)]
    [InlineData("PDF", SaveFormat.Pdf)]
    public void GetWordSaveFormat_SupportedFormats_ReturnsCorrectFormat(string format,
        SaveFormat expected)
    {
        var result = DocumentConverter.GetWordSaveFormat(format);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("xlsx")]
    [InlineData("pptx")]
    [InlineData("png")]
    [InlineData("unknown")]
    public void GetWordSaveFormat_UnsupportedFormats_ThrowsArgumentException(string format)
    {
        var ex = Assert.Throws<ArgumentException>(() => DocumentConverter.GetWordSaveFormat(format));

        Assert.Contains("Unsupported output format for Word", ex.Message);
    }

    #endregion

    #region GetExcelSaveFormat Tests

    [Theory]
    [InlineData("pdf", Aspose.Cells.SaveFormat.Pdf)]
    [InlineData("xlsx", Aspose.Cells.SaveFormat.Xlsx)]
    [InlineData("xls", Aspose.Cells.SaveFormat.Excel97To2003)]
    [InlineData("csv", Aspose.Cells.SaveFormat.Csv)]
    [InlineData("html", Aspose.Cells.SaveFormat.Html)]
    [InlineData("ods", Aspose.Cells.SaveFormat.Ods)]
    [InlineData(".pdf", Aspose.Cells.SaveFormat.Pdf)]
    [InlineData("PDF", Aspose.Cells.SaveFormat.Pdf)]
    public void GetExcelSaveFormat_SupportedFormats_ReturnsCorrectFormat(string format,
        Aspose.Cells.SaveFormat expected)
    {
        var result = DocumentConverter.GetExcelSaveFormat(format);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("docx")]
    [InlineData("pptx")]
    [InlineData("png")]
    [InlineData("unknown")]
    public void GetExcelSaveFormat_UnsupportedFormats_ThrowsArgumentException(string format)
    {
        var ex = Assert.Throws<ArgumentException>(() => DocumentConverter.GetExcelSaveFormat(format));

        Assert.Contains("Unsupported output format for Excel", ex.Message);
    }

    #endregion

    #region GetPresentationSaveFormat Tests

    [Theory]
    [InlineData("pdf", Aspose.Slides.Export.SaveFormat.Pdf)]
    [InlineData("pptx", Aspose.Slides.Export.SaveFormat.Pptx)]
    [InlineData("ppt", Aspose.Slides.Export.SaveFormat.Ppt)]
    [InlineData("html", Aspose.Slides.Export.SaveFormat.Html)]
    [InlineData("odp", Aspose.Slides.Export.SaveFormat.Odp)]
    [InlineData(".pdf", Aspose.Slides.Export.SaveFormat.Pdf)]
    [InlineData("PDF", Aspose.Slides.Export.SaveFormat.Pdf)]
    public void GetPresentationSaveFormat_SupportedFormats_ReturnsCorrectFormat(string format,
        Aspose.Slides.Export.SaveFormat expected)
    {
        var result = DocumentConverter.GetPresentationSaveFormat(format);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("docx")]
    [InlineData("xlsx")]
    [InlineData("png")]
    [InlineData("unknown")]
    public void GetPresentationSaveFormat_UnsupportedFormats_ThrowsArgumentException(string format)
    {
        var ex = Assert.Throws<ArgumentException>(() => DocumentConverter.GetPresentationSaveFormat(format));

        Assert.Contains("Unsupported output format for PowerPoint", ex.Message);
    }

    #endregion

    #region GetPdfSaveFormat Tests

    [Theory]
    [InlineData("docx", Aspose.Pdf.SaveFormat.DocX)]
    [InlineData("doc", Aspose.Pdf.SaveFormat.Doc)]
    [InlineData("html", Aspose.Pdf.SaveFormat.Html)]
    [InlineData("xlsx", Aspose.Pdf.SaveFormat.Excel)]
    [InlineData("pptx", Aspose.Pdf.SaveFormat.Pptx)]
    [InlineData("txt", Aspose.Pdf.SaveFormat.TeX)]
    [InlineData("epub", Aspose.Pdf.SaveFormat.Epub)]
    [InlineData("svg", Aspose.Pdf.SaveFormat.Svg)]
    [InlineData("xps", Aspose.Pdf.SaveFormat.Xps)]
    [InlineData("xml", Aspose.Pdf.SaveFormat.Xml)]
    [InlineData(".docx", Aspose.Pdf.SaveFormat.DocX)]
    [InlineData("DOCX", Aspose.Pdf.SaveFormat.DocX)]
    public void GetPdfSaveFormat_SupportedFormats_ReturnsCorrectFormat(string format, Aspose.Pdf.SaveFormat expected)
    {
        var result = DocumentConverter.GetPdfSaveFormat(format);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("pdf")]
    [InlineData("png")]
    [InlineData("jpg")]
    [InlineData("unknown")]
    public void GetPdfSaveFormat_UnsupportedFormats_ThrowsArgumentException(string format)
    {
        var ex = Assert.Throws<ArgumentException>(() => DocumentConverter.GetPdfSaveFormat(format));

        Assert.Contains("Unsupported output format for PDF", ex.Message);
    }

    #endregion

    #region GetSupportedFormats Edge Case Tests

    [Fact]
    public void GetSupportedFormats_UnknownDocumentType_ReturnsEmpty()
    {
        var formats = DocumentConverter.GetSupportedFormats((DocumentType)999).ToList();

        Assert.Empty(formats);
    }

    [Fact]
    public void IsFormatSupported_UnknownDocumentType_ReturnsFalse()
    {
        var result = DocumentConverter.IsFormatSupported((DocumentType)999, "pdf");

        Assert.False(result);
    }

    #endregion

    #region Word Conversion Tests

    [Fact]
    public void ConvertToStream_WordToPdf_ReturnsValidStream()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Test content for PDF conversion");

        using var stream = DocumentConverter.ConvertToStream(doc, DocumentType.Word, "pdf");

        Assert.NotNull(stream);
        Assert.True(stream.Length > 0);
    }

    [Fact]
    public void ConvertToStream_WordToHtml_ReturnsValidStream()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Test content for HTML conversion");

        using var stream = DocumentConverter.ConvertToStream(doc, DocumentType.Word, "html");

        Assert.NotNull(stream);
        Assert.True(stream.Length > 0);
    }

    [Fact]
    public void ConvertToStream_WordToDocx_ReturnsValidStream()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Test content");

        using var stream = DocumentConverter.ConvertToStream(doc, DocumentType.Word, "docx");

        Assert.NotNull(stream);
        Assert.True(stream.Length > 0);
    }

    [Fact]
    public void ConvertToStream_WordToRtf_ReturnsValidStream()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Test content");

        using var stream = DocumentConverter.ConvertToStream(doc, DocumentType.Word, "rtf");

        Assert.NotNull(stream);
        Assert.True(stream.Length > 0);
    }

    [Fact]
    public void ConvertToStream_WordToTxt_ReturnsValidStream()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Test content");

        using var stream = DocumentConverter.ConvertToStream(doc, DocumentType.Word, "txt");

        Assert.NotNull(stream);
        Assert.True(stream.Length > 0);
    }

    [Fact]
    public void ConvertToBytes_WordToPdf_ReturnsValidBytes()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Test content for PDF conversion");

        var bytes = DocumentConverter.ConvertToBytes(doc, DocumentType.Word, "pdf");

        Assert.NotNull(bytes);
        Assert.True(bytes.Length > 0);
    }

    [SkippableFact]
    public void ConvertWordDocument_ToPdf_CreatesFile()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words);

        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Test content for PDF conversion");
        var outputPath = CreateTestFilePath("output.pdf");

        DocumentConverter.ConvertWordDocument(doc, outputPath, ".pdf", null, new ConversionOptions());

        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    [SkippableFact]
    public void ConvertWordDocument_ToHtml_WithEmbedImages_CreatesFile()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words);

        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Test content for HTML conversion");
        var outputPath = CreateTestFilePath("output.html");
        var options = new ConversionOptions { HtmlEmbedImages = true, HtmlSingleFile = true };

        DocumentConverter.ConvertWordDocument(doc, outputPath, ".html", null, options);

        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    [SkippableFact]
    public void ConvertWordToImages_SinglePage_CreatesFile()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words);

        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Test content for image conversion");
        var outputPath = CreateTestFilePath("output.png");
        var options = new ConversionOptions { PageIndex = 1, Dpi = 150 };

        DocumentConverter.ConvertWordToImages(doc, outputPath, ".png", options);

        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    #endregion

    #region Excel Conversion Tests

    [Fact]
    public void ConvertToStream_ExcelToPdf_ReturnsValidStream()
    {
        using var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Test";

        using var stream = DocumentConverter.ConvertToStream(workbook, DocumentType.Excel, "pdf");

        Assert.NotNull(stream);
        Assert.True(stream.Length > 0);
    }

    [Fact]
    public void ConvertToStream_ExcelToHtml_ReturnsValidStream()
    {
        using var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Test";

        using var stream = DocumentConverter.ConvertToStream(workbook, DocumentType.Excel, "html");

        Assert.NotNull(stream);
        Assert.True(stream.Length > 0);
    }

    [Fact]
    public void ConvertToStream_ExcelToCsv_ReturnsValidStream()
    {
        using var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Test";

        using var stream = DocumentConverter.ConvertToStream(workbook, DocumentType.Excel, "csv");

        Assert.NotNull(stream);
        Assert.True(stream.Length > 0);
    }

    [Fact]
    public void ConvertToStream_ExcelToXlsx_ReturnsValidStream()
    {
        using var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Test";

        using var stream = DocumentConverter.ConvertToStream(workbook, DocumentType.Excel, "xlsx");

        Assert.NotNull(stream);
        Assert.True(stream.Length > 0);
    }

    [Fact]
    public void ConvertToBytes_ExcelToPdf_ReturnsValidBytes()
    {
        using var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Test";

        var bytes = DocumentConverter.ConvertToBytes(workbook, DocumentType.Excel, "pdf");

        Assert.NotNull(bytes);
        Assert.True(bytes.Length > 0);
    }

    [SkippableFact]
    public void ConvertExcelDocument_ToPdf_CreatesFile()
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells);

        using var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Test content";
        var outputPath = CreateTestFilePath("output.pdf");

        DocumentConverter.ConvertExcelDocument(workbook, outputPath, ".pdf", null, new ConversionOptions());

        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    [SkippableFact]
    public void ConvertExcelDocument_ToCsv_CreatesFile()
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells);

        using var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Worksheets[0].Cells["B1"].Value = "Data";
        var outputPath = CreateTestFilePath("output.csv");
        var options = new ConversionOptions { CsvSeparator = "," };

        DocumentConverter.ConvertExcelDocument(workbook, outputPath, ".csv", null, options);

        Assert.True(File.Exists(outputPath));
        var content = File.ReadAllText(outputPath);
        Assert.Contains("Test", content);
    }

    [SkippableFact]
    public void ConvertExcelToImages_SingleSheet_CreatesFile()
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells);

        using var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        var outputPath = CreateTestFilePath("output.png");
        var options = new ConversionOptions { PageIndex = 1, Dpi = 150 };

        DocumentConverter.ConvertExcelToImages(workbook, outputPath, ".png", options);

        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    #endregion

    #region PowerPoint Conversion Tests

    [Fact]
    public void ConvertToStream_PowerPointToPdf_ReturnsValidStream()
    {
        using var presentation = new Presentation();
        _ = presentation.Slides[0];

        using var stream = DocumentConverter.ConvertToStream(presentation, DocumentType.PowerPoint, "pdf");

        Assert.NotNull(stream);
        Assert.True(stream.Length > 0);
    }

    [Fact]
    public void ConvertToStream_PowerPointToHtml_ReturnsValidStream()
    {
        using var presentation = new Presentation();

        using var stream = DocumentConverter.ConvertToStream(presentation, DocumentType.PowerPoint, "html");

        Assert.NotNull(stream);
        Assert.True(stream.Length > 0);
    }

    [Fact]
    public void ConvertToStream_PowerPointToPptx_ReturnsValidStream()
    {
        using var presentation = new Presentation();

        using var stream = DocumentConverter.ConvertToStream(presentation, DocumentType.PowerPoint, "pptx");

        Assert.NotNull(stream);
        Assert.True(stream.Length > 0);
    }

    [Fact]
    public void ConvertToBytes_PowerPointToPdf_ReturnsValidBytes()
    {
        using var presentation = new Presentation();

        var bytes = DocumentConverter.ConvertToBytes(presentation, DocumentType.PowerPoint, "pdf");

        Assert.NotNull(bytes);
        Assert.True(bytes.Length > 0);
    }

    [SkippableFact]
    public void ConvertPowerPointDocument_ToPdf_CreatesFile()
    {
        SkipInEvaluationMode();

        using var presentation = new Presentation();
        var outputPath = CreateTestFilePath("output.pdf");

        DocumentConverter.ConvertPowerPointDocument(presentation, outputPath, ".pdf", null, new ConversionOptions());

        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    [SkippableFact]
    public void ConvertPowerPointDocument_ToHtml_CreatesFile()
    {
        SkipInEvaluationMode();

        using var presentation = new Presentation();
        var outputPath = CreateTestFilePath("output.html");

        DocumentConverter.ConvertPowerPointDocument(presentation, outputPath, ".html", null, new ConversionOptions());

        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    #endregion

    #region PDF Conversion Tests

    [Fact]
    public void ConvertToStream_PdfToDocx_ReturnsValidStream()
    {
        var pdfDoc = new Aspose.Pdf.Document();
        var page = pdfDoc.Pages.Add();
        page.Paragraphs.Add(new TextFragment("Test content"));

        using var stream = DocumentConverter.ConvertToStream(pdfDoc, DocumentType.Pdf, "docx");

        Assert.NotNull(stream);
        Assert.True(stream.Length > 0);
    }

    [Fact]
    public void ConvertToStream_PdfToHtml_ReturnsValidStream()
    {
        var pdfDoc = new Aspose.Pdf.Document();
        var page = pdfDoc.Pages.Add();
        page.Paragraphs.Add(new TextFragment("Test content"));

        using var stream = DocumentConverter.ConvertToStream(pdfDoc, DocumentType.Pdf, "html");

        Assert.NotNull(stream);
        Assert.True(stream.Length > 0);
    }

    [Fact]
    public void ConvertToBytes_PdfToDocx_ReturnsValidBytes()
    {
        var pdfDoc = new Aspose.Pdf.Document();
        var page = pdfDoc.Pages.Add();
        page.Paragraphs.Add(new TextFragment("Test content"));

        var bytes = DocumentConverter.ConvertToBytes(pdfDoc, DocumentType.Pdf, "docx");

        Assert.NotNull(bytes);
        Assert.True(bytes.Length > 0);
    }

    [SkippableFact]
    public void ConvertPdfDocument_ToDocx_CreatesFile()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);

        var pdfDoc = new Aspose.Pdf.Document();
        var page = pdfDoc.Pages.Add();
        page.Paragraphs.Add(new TextFragment("Test content for DOCX conversion"));
        var outputPath = CreateTestFilePath("output.docx");

        DocumentConverter.ConvertPdfDocument(pdfDoc, outputPath, ".docx", new ConversionOptions());

        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    [SkippableFact]
    public void ConvertPdfDocument_ToHtml_CreatesFile()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);

        var pdfDoc = new Aspose.Pdf.Document();
        var page = pdfDoc.Pages.Add();
        page.Paragraphs.Add(new TextFragment("Test content for HTML conversion"));
        var outputPath = CreateTestFilePath("output.html");

        DocumentConverter.ConvertPdfDocument(pdfDoc, outputPath, ".html", new ConversionOptions());

        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    [SkippableFact]
    public void ConvertPdfToImages_SinglePage_CreatesFile()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);

        var pdfDoc = new Aspose.Pdf.Document();
        pdfDoc.Pages.Add();
        var inputPath = CreateTestFilePath("input.pdf");
        pdfDoc.Save(inputPath);
        var outputPath = CreateTestFilePath("output.png");
        var options = new ConversionOptions { PageIndex = 1, Dpi = 150 };

        DocumentConverter.ConvertPdfToImages(inputPath, outputPath, ".png", 1, options);

        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    #endregion

    #region ConversionOptions Tests

    [Fact]
    public void ConvertToStream_WithOptions_AppliesJpegQuality()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Test content");
        var options = new ConversionOptions { JpegQuality = 50 };

        using var stream = DocumentConverter.ConvertToStream(doc, DocumentType.Word, "pdf", options);

        Assert.NotNull(stream);
        Assert.True(stream.Length > 0);
    }

    [Fact]
    public void ConvertToStream_WithOptions_AppliesDpi()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Test content");
        var options = new ConversionOptions { Dpi = 300 };

        using var stream = DocumentConverter.ConvertToStream(doc, DocumentType.Word, "pdf", options);

        Assert.NotNull(stream);
        Assert.True(stream.Length > 0);
    }

    #endregion

    #region Invalid Document Type Tests

    [Fact]
    public void ConvertToStream_InvalidDocumentType_ThrowsArgumentException()
    {
        var doc = new Document();

        var ex = Assert.Throws<ArgumentException>(() =>
            DocumentConverter.ConvertToStream(doc, (DocumentType)999, "pdf"));

        Assert.Contains("not supported", ex.Message);
    }

    [Fact]
    public void ConvertToBytes_InvalidDocumentType_ThrowsArgumentException()
    {
        var doc = new Document();

        var ex = Assert.Throws<ArgumentException>(() =>
            DocumentConverter.ConvertToBytes(doc, (DocumentType)999, "pdf"));

        Assert.Contains("not supported", ex.Message);
    }

    #endregion
}
