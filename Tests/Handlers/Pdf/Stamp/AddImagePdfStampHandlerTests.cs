using AsposeMcpServer.Handlers.Pdf.Stamp;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Stamp;

/// <summary>
///     Tests for <see cref="AddImagePdfStampHandler" />.
///     Validates image stamp creation with various parameters and error handling.
/// </summary>
public class AddImagePdfStampHandlerTests : PdfHandlerTestBase
{
    private readonly AddImagePdfStampHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_AddImage()
    {
        Assert.Equal("add_image", _handler.Operation);
    }

    #endregion

    #region Modification Tracking

    [Fact]
    public void Execute_MarksContextAsModified()
    {
        var imagePath = CreateTempImageFile();
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", imagePath }
        });

        _handler.Execute(context, parameters);

        AssertModified(context);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsImageStampToAllPages()
    {
        var imagePath = CreateTempImageFile();
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", imagePath },
            { "pageIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("all pages", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_AddsImageStampToSpecificPage()
    {
        var imagePath = CreateTempImageFile();
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", imagePath },
            { "pageIndex", 2 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("page 2", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithCustomDimensions_AddsStamp()
    {
        var imagePath = CreateTempImageFile();
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", imagePath },
            { "width", 200.0 },
            { "height", 150.0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithOpacityAndRotation_AddsStamp()
    {
        var imagePath = CreateTempImageFile();
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", imagePath },
            { "opacity", 0.5 },
            { "rotation", 90.0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithMissingImagePath_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", "" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("imagePath", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithNonExistentImagePath_ThrowsFileNotFoundException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", "/nonexistent/path/image.png" }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidPageIndex_ThrowsArgumentException()
    {
        var imagePath = CreateTempImageFile();
        var doc = CreateDocumentWithPages(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", imagePath },
            { "pageIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
