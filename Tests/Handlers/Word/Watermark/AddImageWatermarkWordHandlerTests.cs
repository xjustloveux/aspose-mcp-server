using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Watermark;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Watermark;

public class AddImageWatermarkWordHandlerTests : WordHandlerTestBase
{
    private readonly AddImageWatermarkWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_AddImage()
    {
        Assert.Equal("add_image", _handler.Operation);
    }

    #endregion

    #region Overwrite Existing Watermark

    [Fact]
    public void Execute_OverwritesExistingTextWatermark()
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateEmptyDocument();
        doc.Watermark.SetText("OLD TEXT");
        Assert.Equal(WatermarkType.Text, doc.Watermark.Type);

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", tempFile }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("watermark added", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(WatermarkType.Image, doc.Watermark.Type);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsImageWatermark()
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

        Assert.Contains("watermark added", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(WatermarkType.Image, doc.Watermark.Type);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsFileName()
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

        Assert.Contains(Path.GetFileName(tempFile), result.Message);
    }

    [Fact]
    public void Execute_WithScale_SetsScale()
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", tempFile },
            { "scale", 0.5 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Scale: 0.5", result.Message);
        Assert.Equal(WatermarkType.Image, doc.Watermark.Type);
    }

    [Fact]
    public void Execute_WithIsWashout_SetsWashout()
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", tempFile },
            { "isWashout", false }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Washout: False", result.Message);
        Assert.Equal(WatermarkType.Image, doc.Watermark.Type);
    }

    [Fact]
    public void Execute_WithAllOptions_SetsAllOptions()
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", tempFile },
            { "scale", 2.0 },
            { "isWashout", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Scale: 2", result.Message);
        Assert.Contains("Washout: True", result.Message);
        Assert.Equal(WatermarkType.Image, doc.Watermark.Type);
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
        Assert.Contains("imagePath", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithEmptyImagePath_ThrowsArgumentException()
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
