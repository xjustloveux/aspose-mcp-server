using AsposeMcpServer.Handlers.Pdf.Page;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Page;

/// <summary>
///     Unit tests for ResizePdfPageHandler class.
/// </summary>
public class ResizePdfPageHandlerTests : PdfHandlerTestBase
{
    private readonly ResizePdfPageHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Resize()
    {
        Assert.Equal("resize", _handler.Operation);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsMessageWithDimensions()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "width", 595.0 },
            { "height", 842.0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("595", result.Message);
        Assert.Contains("842", result.Message);
        Assert.Contains("points", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Basic Resize Operations

    [Fact]
    public void Execute_ResizesPage_SetsMediaBoxAndCropBox()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "width", 595.0 },
            { "height", 842.0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("resized", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);

        var page = doc.Pages[1];
        Assert.Equal(0, page.MediaBox.LLX, 1);
        Assert.Equal(0, page.MediaBox.LLY, 1);
        Assert.Equal(595.0, page.MediaBox.URX, 1);
        Assert.Equal(842.0, page.MediaBox.URY, 1);

        Assert.Equal(0, page.CropBox.LLX, 1);
        Assert.Equal(0, page.CropBox.LLY, 1);
        Assert.Equal(595.0, page.CropBox.URX, 1);
        Assert.Equal(842.0, page.CropBox.URY, 1);
    }

    [Theory]
    [InlineData(612.0, 792.0)]
    [InlineData(595.0, 842.0)]
    [InlineData(841.0, 1190.0)]
    public void Execute_WithVariousSizes_SetsCorrectDimensions(double width, double height)
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "width", width },
            { "height", height }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);

        var page = doc.Pages[1];
        Assert.Equal(width, page.MediaBox.URX, 1);
        Assert.Equal(height, page.MediaBox.URY, 1);
    }

    [Fact]
    public void Execute_OnMultiPageDocument_ResizesSpecificPage()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 2 },
            { "width", 400.0 },
            { "height", 600.0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("2", result.Message);
        AssertModified(context);

        Assert.Equal(400.0, doc.Pages[2].MediaBox.URX, 1);
        Assert.Equal(600.0, doc.Pages[2].MediaBox.URY, 1);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidPageIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 99 },
            { "width", 595.0 },
            { "height", 842.0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("pageIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithZeroPageIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 0 },
            { "width", 595.0 },
            { "height", 842.0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("pageIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithNonPositiveWidth_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "width", 0.0 },
            { "height", 842.0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("positive", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithNegativeHeight_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "width", 595.0 },
            { "height", -100.0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("positive", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutRequiredParameters_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 }
        });

        Assert.ThrowsAny<Exception>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
