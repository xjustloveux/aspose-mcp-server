using AsposeMcpServer.Handlers.Word.HeaderFooter;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.HeaderFooter;

public class SetHeaderImageHandlerTests : WordHandlerTestBase
{
    private readonly SetHeaderImageHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetHeaderImage()
    {
        Assert.Equal("set_header_image", _handler.Operation);
    }

    #endregion

    #region Basic Operations

    [Fact]
    public void Execute_SetsHeaderImage()
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", tempFile }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("header image set", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithAlignment_SetsAlignment()
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", tempFile },
            { "alignment", "center" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("header image set", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithFloating_SetsFloatingImage()
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", tempFile },
            { "isFloating", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("floating", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithDimensions_SetsDimensions()
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", tempFile },
            { "imageWidth", 100.0 },
            { "imageHeight", 50.0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("header image set", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutImagePath_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("imagePath", ex.Message);
    }

    [Fact]
    public void Execute_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", "/nonexistent/image.png" }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
