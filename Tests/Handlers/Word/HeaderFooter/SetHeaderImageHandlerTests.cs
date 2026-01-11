using AsposeMcpServer.Handlers.Word.HeaderFooter;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("header image set", result.ToLower());
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("header image set", result.ToLower());
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("floating", result.ToLower());
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("header image set", result.ToLower());
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
