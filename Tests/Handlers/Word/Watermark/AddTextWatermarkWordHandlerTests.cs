using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Watermark;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Watermark;

public class AddTextWatermarkWordHandlerTests : WordHandlerTestBase
{
    private readonly AddTextWatermarkWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Overwrite Existing Watermark

    [Fact]
    public void Execute_OverwritesExistingWatermark()
    {
        var doc = CreateEmptyDocument();
        doc.Watermark.SetText("OLD");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "NEW" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("watermark added", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(WatermarkType.Text, doc.Watermark.Type);
        AssertModified(context);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsTextWatermark()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "CONFIDENTIAL" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("watermark added", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(WatermarkType.Text, doc.Watermark.Type);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithFontFamily_SetsFontFamily()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "DRAFT" },
            { "fontFamily", "Times New Roman" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("watermark added", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(WatermarkType.Text, doc.Watermark.Type);
    }

    [Fact]
    public void Execute_WithFontSize_SetsFontSize()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "SAMPLE" },
            { "fontSize", 48.0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("watermark added", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(WatermarkType.Text, doc.Watermark.Type);
    }

    [Fact]
    public void Execute_WithLayoutDiagonal_SetsLayout()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "DIAGONAL" },
            { "layout", "Diagonal" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("watermark added", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(WatermarkType.Text, doc.Watermark.Type);
    }

    [Fact]
    public void Execute_WithLayoutHorizontal_SetsLayout()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "HORIZONTAL" },
            { "layout", "Horizontal" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("watermark added", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(WatermarkType.Text, doc.Watermark.Type);
    }

    [Fact]
    public void Execute_WithSemitransparent_SetsSemitransparent()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "TRANSPARENT" },
            { "isSemitransparent", false }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("watermark added", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(WatermarkType.Text, doc.Watermark.Type);
    }

    [Fact]
    public void Execute_WithAllOptions_SetsAllOptions()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "COMPLETE" },
            { "fontFamily", "Verdana" },
            { "fontSize", 60.0 },
            { "isSemitransparent", true },
            { "layout", "Horizontal" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("watermark added", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(WatermarkType.Text, doc.Watermark.Type);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutText_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("text", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithEmptyText_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("text", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
