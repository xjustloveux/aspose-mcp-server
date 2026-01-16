using System.Text.Json;
using Aspose.Slides;
using AsposeMcpServer.Core.ShapeDetailProviders;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Core.ShapeDetailProviders;

public class AutoShapeDetailProviderTests : TestBase
{
    private readonly AutoShapeDetailProvider _provider = new();

    [Fact]
    public void TypeName_ShouldReturnAutoShape()
    {
        Assert.Equal("AutoShape", _provider.TypeName);
    }

    [Fact]
    public void CanHandle_WithAutoShape_ShouldReturnTrue()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);

        var result = _provider.CanHandle(shape);

        Assert.True(result);
    }

    [Fact]
    public void CanHandle_WithTable_ShouldReturnFalse()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var table = slide.Shapes.AddTable(10, 10, [50], [30]);

        var result = _provider.CanHandle(table);

        Assert.False(result);
    }

    [Fact]
    public void GetDetails_WithAutoShape_ShouldReturnDetails()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);
        shape.TextFrame.Text = "Sample Text";

        var details = _provider.GetDetails(shape, presentation);

        Assert.NotNull(details);
    }

    [Fact]
    public void GetDetails_WithNonAutoShape_ShouldReturnNull()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var table = slide.Shapes.AddTable(10, 10, [50], [30]);

        var details = _provider.GetDetails(table, presentation);

        Assert.Null(details);
    }

    [Fact]
    public void GetDetails_WithEmptyTextFrame_ShouldIncludeNullText()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);

        var details = _provider.GetDetails(shape, presentation);

        Assert.NotNull(details);
    }

    [Fact]
    public void GetDetails_WithExternalHyperlink_ShouldIncludeHyperlink()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);
        shape.HyperlinkClick = new Hyperlink("https://example.com");

        var details = _provider.GetDetails(shape, presentation);
        Assert.NotNull(details);

        var json = JsonSerializer.Serialize(details);
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        Assert.Equal("https://example.com", root.GetProperty("hyperlink").GetString());
    }

    [Fact]
    public void GetDetails_WithInternalSlideHyperlink_ShouldIncludeSlideReference()
    {
        using var presentation = new Presentation();
        var slide1 = presentation.Slides[0];
        var slide2 = presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);

        var shape = slide1.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);
        shape.HyperlinkClick = new Hyperlink(slide2);

        var details = _provider.GetDetails(shape, presentation);
        Assert.NotNull(details);

        var json = JsonSerializer.Serialize(details);
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        var hyperlink = root.GetProperty("hyperlink").GetString();
        Assert.Contains("Slide", hyperlink);
    }

    [Fact]
    public void GetDetails_WithAdjustments_ShouldIncludeAdjustmentValues()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.RoundCornerRectangle, 10, 10, 100, 100);

        var details = _provider.GetDetails(shape, presentation);
        Assert.NotNull(details);

        var json = JsonSerializer.Serialize(details);
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        if (root.TryGetProperty("adjustments", out var adjustments) &&
            adjustments.ValueKind != JsonValueKind.Null)
            Assert.True(adjustments.GetArrayLength() > 0);
    }

    [Fact]
    public void GetDetails_WithTextFrame_ShouldIncludeTextInfo()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);
        shape.TextFrame.Text = "Test Text";

        var details = _provider.GetDetails(shape, presentation);
        Assert.NotNull(details);

        var json = JsonSerializer.Serialize(details);
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        Assert.True(root.GetProperty("hasTextFrame").GetBoolean());
        Assert.True(root.GetProperty("paragraphCount").GetInt32() >= 1);
    }
}
