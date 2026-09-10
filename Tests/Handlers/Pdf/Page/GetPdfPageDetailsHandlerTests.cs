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
        Assert.Equal("details", _handler.Operation);
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

        Assert.Equal(595, result.Width);
        Assert.Equal(842, result.Height);
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
        // "x >= 0 || x < 0" holds for every number, so it asserted nothing. A page box
        // describes a rectangle with positive area, and the fixture uses the default
        // A4-sized page, so the corners must be ordered and the box must be non-empty.
        Assert.True(result.MediaBox.Urx > result.MediaBox.Llx, "MediaBox has no width");
        Assert.True(result.MediaBox.Ury > result.MediaBox.Lly, "MediaBox has no height");
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
        // "x >= 0 || x < 0" holds for every number, so it asserted nothing. A page box
        // describes a rectangle with positive area, and the fixture uses the default
        // A4-sized page, so the corners must be ordered and the box must be non-empty.
        Assert.True(result.CropBox.Urx > result.CropBox.Llx, "CropBox has no width");
        Assert.True(result.CropBox.Ury > result.CropBox.Lly, "CropBox has no height");
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

        Assert.Equal(0, result.Annotations);
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

        Assert.Equal(0, result.Paragraphs);
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

        Assert.Equal(0, result.Images);
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
