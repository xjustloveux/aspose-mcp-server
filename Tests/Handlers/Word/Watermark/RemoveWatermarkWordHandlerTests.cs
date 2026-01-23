using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Watermark;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;
using SkiaSharp;

namespace AsposeMcpServer.Tests.Handlers.Word.Watermark;

public class RemoveWatermarkWordHandlerTests : WordHandlerTestBase
{
    private readonly RemoveWatermarkWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Remove()
    {
        Assert.Equal("remove", _handler.Operation);
    }

    #endregion

    #region No Watermark Case

    [Fact]
    public void Execute_WithNoWatermark_ReturnsNoWatermarkMessage()
    {
        var doc = CreateEmptyDocument();
        Assert.Equal(WatermarkType.None, doc.Watermark.Type);

        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("no watermark", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertNotModified(context);
    }

    #endregion

    #region Helper Methods

    private static void AddImageWatermark(Document doc, string imagePath)
    {
        using var bitmap = SKBitmap.Decode(imagePath);
        doc.Watermark.SetImage(bitmap, new ImageWatermarkOptions());
    }

    #endregion

    #region Basic Remove Operations

    [Fact]
    public void Execute_RemovesTextWatermark()
    {
        var doc = CreateEmptyDocument();
        doc.Watermark.SetText("CONFIDENTIAL");
        Assert.Equal(WatermarkType.Text, doc.Watermark.Type);

        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("removed", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(WatermarkType.None, doc.Watermark.Type);
        AssertModified(context);
    }

    [Fact]
    public void Execute_RemovesImageWatermark()
    {
        var tempFile = CreateTempImageFile();
        var doc = CreateEmptyDocument();
        AddImageWatermark(doc, tempFile);
        Assert.Equal(WatermarkType.Image, doc.Watermark.Type);

        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("removed", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(WatermarkType.None, doc.Watermark.Type);
        AssertModified(context);
    }

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_WithWatermark_MarksAsModified()
    {
        var doc = CreateEmptyDocument();
        doc.Watermark.SetText("TEST");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        AssertModified(context);
    }

    [Fact]
    public void Execute_WithoutWatermark_DoesNotModify()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        AssertNotModified(context);
    }

    #endregion
}
