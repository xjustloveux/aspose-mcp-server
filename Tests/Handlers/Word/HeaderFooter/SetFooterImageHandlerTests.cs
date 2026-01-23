using AsposeMcpServer.Handlers.Word.HeaderFooter;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.HeaderFooter;

public class SetFooterImageHandlerTests : WordHandlerTestBase
{
    private readonly SetFooterImageHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetFooterImage()
    {
        Assert.Equal("set_footer_image", _handler.Operation);
    }

    #endregion

    #region Alignment Tests

    [Theory]
    [InlineData("left")]
    [InlineData("center")]
    [InlineData("right")]
    public void Execute_WithVariousAlignments_SetsCorrectAlignment(string alignment)
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", tempFile },
            { "alignment", alignment }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("footer image set", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    #endregion

    #region Header Footer Type Tests

    [Theory]
    [InlineData("primary")]
    [InlineData("first")]
    [InlineData("even")]
    public void Execute_WithHeaderFooterType_SetsCorrectType(string headerFooterType)
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", tempFile },
            { "headerFooterType", headerFooterType }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("footer image set", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    #endregion

    #region Remove Existing Tests

    [Fact]
    public void Execute_WithRemoveExistingFalse_KeepsExistingImages()
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", tempFile },
            { "removeExisting", false }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("footer image set", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    #endregion

    #region Basic Operations

    [Fact]
    public void Execute_SetsFooterImage()
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

        Assert.Contains("footer image set", result.Message, StringComparison.OrdinalIgnoreCase);
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
            { "alignment", "right" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("footer image set", result.Message, StringComparison.OrdinalIgnoreCase);
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

    #region Image Size Tests

    [Fact]
    public void Execute_WithImageWidth_SetsWidth()
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", tempFile },
            { "imageWidth", 100.0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("footer image set", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithImageHeight_SetsHeight()
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", tempFile },
            { "imageHeight", 50.0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("footer image set", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithBothDimensions_SetsBoth()
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", tempFile },
            { "imageWidth", 150.0 },
            { "imageHeight", 75.0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("footer image set", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    #endregion

    #region Floating Image Tests

    [Fact]
    public void Execute_WithIsFloating_CreatesFloatingImage()
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
        AssertModified(context);
    }

    [Theory]
    [InlineData("left")]
    [InlineData("center")]
    [InlineData("right")]
    public void Execute_WithFloatingAndAlignment_PositionsCorrectly(string alignment)
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", tempFile },
            { "isFloating", true },
            { "alignment", alignment }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("floating", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    #endregion
}
