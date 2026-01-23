using AsposeMcpServer.Handlers.Pdf.Text;
using AsposeMcpServer.Results.Pdf.Text;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Text;

public class ExtractPdfTextHandlerTests : PdfHandlerTestBase
{
    private readonly ExtractPdfTextHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Extract()
    {
        Assert.Equal("extract", _handler.Operation);
    }

    #endregion

    #region Text Content

    [Fact]
    public void Execute_ReturnsTextProperty()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<ExtractPdfTextResult>(res);

        Assert.NotNull(result.Text);
    }

    #endregion

    #region Extraction Mode

    [Theory]
    [InlineData("pure")]
    [InlineData("raw")]
    public void Execute_WithExtractionMode_ExtractsText(string mode)
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "extractionMode", mode }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<ExtractPdfTextResult>(res);

        Assert.NotNull(result.Text);
    }

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_DoesNotModifyDocument()
    {
        var doc = CreateDocumentWithPages(3);
        var initialCount = doc.Pages.Count;
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        Assert.Equal(initialCount, doc.Pages.Count);
        AssertNotModified(context);
    }

    #endregion

    #region Error Handling

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(10)]
    public void Execute_WithInvalidPageIndex_ThrowsArgumentException(int invalidIndex)
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", invalidIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("pageIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Basic Extract Operations

    [Fact]
    public void Execute_ExtractsText()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<ExtractPdfTextResult>(res);

        Assert.True(result.PageIndex >= 0);
        Assert.True(result.TotalPages >= 0);
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_ReturnsPageIndex()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 2 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<ExtractPdfTextResult>(res);

        Assert.Equal(2, result.PageIndex);
    }

    [Fact]
    public void Execute_ReturnsTotalPages()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<ExtractPdfTextResult>(res);

        Assert.Equal(3, result.TotalPages);
    }

    #endregion

    #region Page Index

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void Execute_ExtractsFromVariousPages(int pageIndex)
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", pageIndex }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<ExtractPdfTextResult>(res);

        Assert.Equal(pageIndex, result.PageIndex);
    }

    [Fact]
    public void Execute_DefaultPageIndex_ExtractsFromFirstPage()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<ExtractPdfTextResult>(res);

        Assert.Equal(1, result.PageIndex);
    }

    #endregion

    #region Include Font Info

    [Fact]
    public void Execute_WithIncludeFontInfo_ReturnsFragments()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "includeFontInfo", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<ExtractPdfTextResult>(res);

        Assert.NotNull(result.Fragments);
        Assert.NotNull(result.FragmentCount);
    }

    [Fact]
    public void Execute_WithoutIncludeFontInfo_ReturnsText()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "includeFontInfo", false }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<ExtractPdfTextResult>(res);

        Assert.NotNull(result.Text);
        Assert.Null(result.Fragments);
    }

    #endregion
}
