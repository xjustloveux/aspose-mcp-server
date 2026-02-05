using AsposeMcpServer.Handlers.Pdf.HeaderFooter;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.HeaderFooter;

/// <summary>
///     Tests for <see cref="AddImagePdfHeaderFooterHandler" />.
/// </summary>
public class AddImagePdfHeaderFooterHandlerTests : PdfHandlerTestBase
{
    private readonly AddImagePdfHeaderFooterHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_AddImage()
    {
        Assert.Equal("add_image", _handler.Operation);
    }

    #endregion

    #region Page Range

    [Fact]
    public void Execute_WithPageRange_AppliesOnlyToSelectedPages()
    {
        var doc = CreateDocumentWithPages(4);
        var context = CreateContext(doc);
        var imagePath = CreateTempImageFile();
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", imagePath },
            { "pageRange", "2-4" }
        });

        var result = _handler.Execute(context, parameters);

        var successResult = Assert.IsType<SuccessResult>(result);
        Assert.Contains("3 page(s)", successResult.Message);
        AssertModified(context);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_WithValidImagePath_AddsImageStampToAllPages()
    {
        var doc = CreateDocumentWithPages(2);
        var context = CreateContext(doc);
        var imagePath = CreateTempImageFile();
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", imagePath }
        });

        var result = _handler.Execute(context, parameters);

        var successResult = Assert.IsType<SuccessResult>(result);
        Assert.Contains("2 page(s)", successResult.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithSinglePage_AddsImageStamp()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var imagePath = CreateTempImageFile();
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", imagePath }
        });

        var result = _handler.Execute(context, parameters);

        var successResult = Assert.IsType<SuccessResult>(result);
        Assert.Contains("1 page(s)", successResult.Message);
        AssertModified(context);
    }

    #endregion

    #region Position

    [Fact]
    public void Execute_WithHeaderPosition_ReturnsHeaderMessage()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var imagePath = CreateTempImageFile();
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", imagePath },
            { "position", "header" }
        });

        var result = _handler.Execute(context, parameters);

        var successResult = Assert.IsType<SuccessResult>(result);
        Assert.Contains("header", successResult.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithFooterPosition_ReturnsFooterMessage()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var imagePath = CreateTempImageFile();
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", imagePath },
            { "position", "footer" }
        });

        var result = _handler.Execute(context, parameters);

        var successResult = Assert.IsType<SuccessResult>(result);
        Assert.Contains("footer", successResult.Message);
        AssertModified(context);
    }

    #endregion

    #region Custom Width and Height

    [Fact]
    public void Execute_WithCustomWidth_Succeeds()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var imagePath = CreateTempImageFile();
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", imagePath },
            { "width", 100.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithCustomHeight_Succeeds()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var imagePath = CreateTempImageFile();
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", imagePath },
            { "height", 50.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithCustomWidthAndHeight_Succeeds()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var imagePath = CreateTempImageFile();
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", imagePath },
            { "width", 100.0 },
            { "height", 50.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithMissingImagePath_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNonExistentImagePath_ThrowsFileNotFoundException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", Path.Combine(TestDir, "nonexistent.bmp") }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
