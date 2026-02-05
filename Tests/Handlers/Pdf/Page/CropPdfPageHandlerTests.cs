using AsposeMcpServer.Handlers.Pdf.Page;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Page;

/// <summary>
///     Unit tests for CropPdfPageHandler class.
/// </summary>
public class CropPdfPageHandlerTests : PdfHandlerTestBase
{
    private readonly CropPdfPageHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Crop()
    {
        Assert.Equal("crop", _handler.Operation);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsMessageWithPageIndexAndDimensions()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "x", 50.0 },
            { "y", 50.0 },
            { "width", 400.0 },
            { "height", 600.0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("1", result.Message);
        Assert.Contains("400", result.Message);
        Assert.Contains("600", result.Message);
    }

    #endregion

    #region Basic Crop Operations

    [Fact]
    public void Execute_CropsPage_SetsCropBox()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "x", 50.0 },
            { "y", 50.0 },
            { "width", 400.0 },
            { "height", 600.0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("cropped", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);

        var cropBox = doc.Pages[1].CropBox;
        Assert.Equal(50.0, cropBox.LLX, 1);
        Assert.Equal(50.0, cropBox.LLY, 1);
        Assert.Equal(450.0, cropBox.URX, 1);
        Assert.Equal(650.0, cropBox.URY, 1);
    }

    [Theory]
    [InlineData(0.0, 0.0, 200.0, 300.0)]
    [InlineData(100.0, 100.0, 300.0, 400.0)]
    [InlineData(10.0, 20.0, 500.0, 700.0)]
    public void Execute_WithVariousCoordinates_SetsCropBoxCorrectly(double x, double y, double width, double height)
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "x", x },
            { "y", y },
            { "width", width },
            { "height", height }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);

        var cropBox = doc.Pages[1].CropBox;
        Assert.Equal(x, cropBox.LLX, 1);
        Assert.Equal(y, cropBox.LLY, 1);
        Assert.Equal(x + width, cropBox.URX, 1);
        Assert.Equal(y + height, cropBox.URY, 1);
    }

    [Fact]
    public void Execute_OnMultiPageDocument_CropsSpecificPage()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 2 },
            { "x", 10.0 },
            { "y", 20.0 },
            { "width", 300.0 },
            { "height", 400.0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("2", result.Message);
        AssertModified(context);
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
            { "x", 50.0 },
            { "y", 50.0 },
            { "width", 400.0 },
            { "height", 600.0 }
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
            { "x", 50.0 },
            { "y", 50.0 },
            { "width", 400.0 },
            { "height", 600.0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("pageIndex", ex.Message);
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
