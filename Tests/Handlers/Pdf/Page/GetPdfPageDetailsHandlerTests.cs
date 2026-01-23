using AsposeMcpServer.Handlers.Pdf.Page;
using AsposeMcpServer.Results.Pdf.Page;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Page;

public class GetPdfPageDetailsHandlerTests : PdfHandlerTestBase
{
    private readonly GetPdfPageDetailsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetDetails()
    {
        Assert.Equal("get_details", _handler.Operation);
    }

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_DoesNotModifyDocument()
    {
        var doc = CreateDocumentWithPages(3);
        var initialCount = doc.Pages.Count;
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialCount, doc.Pages.Count);
        AssertNotModified(context);
    }

    #endregion

    #region Error Handling - Missing Parameter

    [Fact]
    public void Execute_WithoutPageIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("pageIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Basic Details Retrieval

    [Fact]
    public void Execute_ReturnsPageDetails()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPdfPageDetailsResult>(res);

        Assert.Equal(1, result.PageIndex);
        AssertNotModified(context);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void Execute_ReturnsDetailsForVariousPages(int pageIndex)
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", pageIndex }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPdfPageDetailsResult>(res);

        Assert.Equal(pageIndex, result.PageIndex);
        AssertNotModified(context);
    }

    #endregion

    #region Detail Properties

    [Fact]
    public void Execute_ReturnsDimensions()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPdfPageDetailsResult>(res);

        Assert.True(result.Width >= 0);
        Assert.True(result.Height >= 0);
    }

    [Fact]
    public void Execute_ReturnsRotation()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPdfPageDetailsResult>(res);

        Assert.NotNull(result.Rotation);
    }

    [Fact]
    public void Execute_ReturnsMediaBox()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPdfPageDetailsResult>(res);

        Assert.NotNull(result.MediaBox);
        Assert.True(result.MediaBox.Llx >= 0 || result.MediaBox.Llx < 0);
        Assert.True(result.MediaBox.Lly >= 0 || result.MediaBox.Lly < 0);
        Assert.True(result.MediaBox.Urx >= 0 || result.MediaBox.Urx < 0);
        Assert.True(result.MediaBox.Ury >= 0 || result.MediaBox.Ury < 0);
    }

    [Fact]
    public void Execute_ReturnsCropBox()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPdfPageDetailsResult>(res);

        Assert.NotNull(result.CropBox);
        Assert.True(result.CropBox.Llx >= 0 || result.CropBox.Llx < 0);
        Assert.True(result.CropBox.Lly >= 0 || result.CropBox.Lly < 0);
        Assert.True(result.CropBox.Urx >= 0 || result.CropBox.Urx < 0);
        Assert.True(result.CropBox.Ury >= 0 || result.CropBox.Ury < 0);
    }

    [Fact]
    public void Execute_ReturnsAnnotationsCount()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPdfPageDetailsResult>(res);

        Assert.True(result.Annotations >= 0);
    }

    [Fact]
    public void Execute_ReturnsParagraphsCount()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPdfPageDetailsResult>(res);

        Assert.True(result.Paragraphs >= 0);
    }

    [Fact]
    public void Execute_ReturnsImagesCount()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPdfPageDetailsResult>(res);

        Assert.True(result.Images >= 0);
    }

    #endregion

    #region Error Handling - Invalid Page Index

    [Theory]
    [InlineData(3, 4)]
    [InlineData(3, 5)]
    [InlineData(3, 100)]
    public void Execute_WithPageIndexOutOfRange_ThrowsArgumentException(int totalPages, int invalidIndex)
    {
        var doc = CreateDocumentWithPages(totalPages);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", invalidIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("pageIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(-5)]
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
}
