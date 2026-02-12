using Aspose.Words;
using AsposeMcpServer.Core.Conversion;
using AsposeMcpServer.Core.Session;

namespace AsposeMcpServer.Tests.Core.Conversion;

/// <summary>
///     Unit tests for DocumentConversionService class.
/// </summary>
public class DocumentConversionServiceTests
{
    private readonly DocumentConversionService _service = new();

    #region ConvertToBytes Validation Tests

    [Fact]
    public void ConvertToBytes_NullDocument_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>(() =>
            _service.ConvertToBytes(null!, DocumentType.Word, "pdf"));
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
        var mimeType = _service.GetMimeType(format);

        Assert.Equal(expectedMimeType, mimeType);
    }

    [Theory]
    [InlineData("unknown")]
    [InlineData("xyz")]
    [InlineData("")]
    public void GetMimeType_UnknownFormats_ReturnsOctetStream(string format)
    {
        var mimeType = _service.GetMimeType(format);

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
        var result = _service.IsFormatSupported(docType, format);

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
        var result = _service.IsFormatSupported(docType, format);

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
        var result = _service.IsFormatSupported(docType, format);

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
        var result = _service.IsFormatSupported(docType, format);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("PDF")]
    [InlineData("Pdf")]
    [InlineData(".pdf")]
    [InlineData(".PDF")]
    public void IsFormatSupported_CaseInsensitive_ReturnsTrue(string format)
    {
        var result = _service.IsFormatSupported(DocumentType.Word, format);

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
        var result = DocumentConversionService.IsExcelImageFormat(format);

        Assert.Equal(expected, result);
    }

    #endregion

    #region GetSupportedFormats Tests

    [Fact]
    public void GetSupportedFormats_Word_ReturnsExpectedFormats()
    {
        var formats = _service.GetSupportedFormats(DocumentType.Word).ToList();

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
        var formats = _service.GetSupportedFormats(DocumentType.Excel).ToList();

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
        var formats = _service.GetSupportedFormats(DocumentType.PowerPoint).ToList();

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
        var formats = _service.GetSupportedFormats(DocumentType.Pdf).ToList();

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
            _service.ConvertToStream(null!, DocumentType.Word, "pdf"));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void ConvertToStream_InvalidFormat_ThrowsArgumentException(string? format)
    {
        var doc = new Document();

        Assert.Throws<ArgumentException>(() =>
            _service.ConvertToStream(doc, DocumentType.Word, format!));
    }

    [Fact]
    public void ConvertToStream_UnsupportedFormat_ThrowsArgumentException()
    {
        var doc = new Document();

        var ex = Assert.Throws<ArgumentException>(() =>
            _service.ConvertToStream(doc, DocumentType.Word, "xlsx"));

        Assert.Contains("not supported", ex.Message);
    }

    #endregion
}
