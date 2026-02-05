using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Handlers.Pdf.Stamp;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Stamp;

/// <summary>
///     Tests for <see cref="AddPdfStampHandler" />.
///     Validates PDF page stamp creation with various parameters and error handling.
/// </summary>
public class AddPdfStampHandlerTests : PdfHandlerTestBase
{
    private readonly AddPdfStampHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_AddPdf()
    {
        Assert.Equal("add_pdf", _handler.Operation);
    }

    #endregion

    #region Modification Tracking

    [Fact]
    public void Execute_MarksContextAsModified()
    {
        var stampPath = CreateStampPdf();
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pdfPath", stampPath }
        });

        _handler.Execute(context, parameters);

        AssertModified(context);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsPdfStampToAllPages()
    {
        var stampPath = CreateStampPdf();
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pdfPath", stampPath },
            { "pageIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("all pages", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_AddsPdfStampToSpecificPage()
    {
        var stampPath = CreateStampPdf();
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pdfPath", stampPath },
            { "pageIndex", 2 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("page 2", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithCustomStampPageIndex_AddsStamp()
    {
        var stampPath = CreateMultiPageStampPdf(3);
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pdfPath", stampPath },
            { "stampPageIndex", 2 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("stamp page: 2", result.Message);
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithMissingPdfPath_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pdfPath", "" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("pdfPath", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithNonExistentPdfPath_ThrowsFileNotFoundException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pdfPath", "/nonexistent/path/stamp.pdf" }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidPageIndex_ThrowsArgumentException()
    {
        var stampPath = CreateStampPdf();
        var doc = CreateDocumentWithPages(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pdfPath", stampPath },
            { "pageIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Helper Methods

    private string CreateStampPdf()
    {
        var stampDoc = CreateDocumentWithText("Stamp Content");
        var stampPath = Path.Combine(TestDir, $"stamp_{Guid.NewGuid()}.pdf");
        stampDoc.Save(stampPath);
        stampDoc.Dispose();
        return stampPath;
    }

    private string CreateMultiPageStampPdf(int pageCount)
    {
        var stampDoc = new Document();
        for (var i = 0; i < pageCount; i++)
        {
            var page = stampDoc.Pages.Add();
            page.Paragraphs.Add(new TextFragment($"Stamp Page {i + 1}"));
        }

        var stampPath = Path.Combine(TestDir, $"stamp_multi_{Guid.NewGuid()}.pdf");
        stampDoc.Save(stampPath);
        stampDoc.Dispose();
        return stampPath;
    }

    #endregion
}
