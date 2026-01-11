using System.Text.Json;
using AsposeMcpServer.Handlers.Pdf.Page;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(1, json.RootElement.GetProperty("pageIndex").GetInt32());
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

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(pageIndex, json.RootElement.GetProperty("pageIndex").GetInt32());
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

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("width", out _));
        Assert.True(json.RootElement.TryGetProperty("height", out _));
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

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("rotation", out _));
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

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("mediaBox", out var mediaBox));
        Assert.True(mediaBox.TryGetProperty("llx", out _));
        Assert.True(mediaBox.TryGetProperty("lly", out _));
        Assert.True(mediaBox.TryGetProperty("urx", out _));
        Assert.True(mediaBox.TryGetProperty("ury", out _));
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

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("cropBox", out var cropBox));
        Assert.True(cropBox.TryGetProperty("llx", out _));
        Assert.True(cropBox.TryGetProperty("lly", out _));
        Assert.True(cropBox.TryGetProperty("urx", out _));
        Assert.True(cropBox.TryGetProperty("ury", out _));
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

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("annotations", out var annotations));
        Assert.True(annotations.TryGetInt32(out _));
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

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("paragraphs", out var paragraphs));
        Assert.True(paragraphs.TryGetInt32(out _));
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

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("images", out var images));
        Assert.True(images.TryGetInt32(out _));
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
