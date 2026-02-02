using System.Drawing;
using Aspose.Slides;
using AsposeMcpServer.Core.ShapeDetailProviders.Details;
using AsposeMcpServer.Core.ShapeDetailProviders.Providers;
using AsposeMcpServer.Tests.Infrastructure;

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

    [SkippableFact]
    public void GetDetails_WithAutoShape_ShouldReturnDetails()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Text content is truncated in evaluation mode");
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);
        shape.TextFrame.Text = "Sample Text";

        var details = _provider.GetDetails(shape, presentation);
        var autoShape = Assert.IsType<AutoShapeDetails>(details);

        Assert.Equal("Rectangle", autoShape.ShapeType);
        Assert.Equal("Sample Text", autoShape.Text);
        Assert.True(autoShape.HasTextFrame);
        Assert.True(autoShape.ParagraphCount >= 1);
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
        var autoShape = Assert.IsType<AutoShapeDetails>(details);

        Assert.Equal("Rectangle", autoShape.ShapeType);
        Assert.True(autoShape.HasTextFrame);
    }

    [Fact]
    public void GetDetails_WithExternalHyperlink_ShouldIncludeHyperlink()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);
        shape.HyperlinkClick = new Hyperlink("https://example.com");

        var details = _provider.GetDetails(shape, presentation);
        var autoShape = Assert.IsType<AutoShapeDetails>(details);

        Assert.Equal("https://example.com", autoShape.Hyperlink);
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
        var autoShape = Assert.IsType<AutoShapeDetails>(details);

        Assert.NotNull(autoShape.Hyperlink);
        Assert.Contains("Slide", autoShape.Hyperlink);
    }

    [Fact]
    public void GetDetails_WithAdjustments_ShouldIncludeAdjustmentValues()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.RoundCornerRectangle, 10, 10, 100, 100);

        var details = _provider.GetDetails(shape, presentation);
        var autoShape = Assert.IsType<AutoShapeDetails>(details);

        if (autoShape.Adjustments != null)
            Assert.True(autoShape.Adjustments.Count > 0);
    }

    [Fact]
    public void GetDetails_WithTextFrame_ShouldIncludeTextInfo()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);
        shape.TextFrame.Text = "Test Text";

        var details = _provider.GetDetails(shape, presentation);
        var autoShape = Assert.IsType<AutoShapeDetails>(details);

        Assert.True(autoShape.HasTextFrame);
        Assert.True(autoShape.ParagraphCount >= 1);
    }

    [Fact]
    public void GetDetails_WithSolidFill_ShouldReturnFillColor()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);
        shape.FillFormat.FillType = FillType.Solid;
        shape.FillFormat.SolidFillColor.Color = Color.FromArgb(255, 0, 128);

        var details = _provider.GetDetails(shape, presentation);
        var autoShape = Assert.IsType<AutoShapeDetails>(details);

        Assert.Equal("#FF0080", autoShape.FillColor);
        Assert.Null(autoShape.Transparency);
    }

    [Fact]
    public void GetDetails_WithSolidFillAndTransparency_ShouldReturnTransparency()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);
        shape.FillFormat.FillType = FillType.Solid;
        shape.FillFormat.SolidFillColor.Color = Color.FromArgb(128, 255, 0, 0);

        var details = _provider.GetDetails(shape, presentation);
        var autoShape = Assert.IsType<AutoShapeDetails>(details);

        Assert.NotNull(autoShape.FillColor);
        Assert.NotNull(autoShape.Transparency);
        Assert.True(autoShape.Transparency > 0);
    }

    [Fact]
    public void GetDetails_WithNoSolidFill_ShouldReturnNullFillColor()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);
        shape.FillFormat.FillType = FillType.NoFill;

        var details = _provider.GetDetails(shape, presentation);
        var autoShape = Assert.IsType<AutoShapeDetails>(details);

        Assert.Null(autoShape.FillColor);
        Assert.Null(autoShape.Transparency);
    }

    [Fact]
    public void GetDetails_WithLineFormat_ShouldReturnLineProperties()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);
        shape.LineFormat.FillFormat.FillType = FillType.Solid;
        shape.LineFormat.FillFormat.SolidFillColor.Color = Color.Red;
        shape.LineFormat.Width = 3.0;
        shape.LineFormat.DashStyle = LineDashStyle.Dash;

        var details = _provider.GetDetails(shape, presentation);
        var autoShape = Assert.IsType<AutoShapeDetails>(details);

        Assert.Equal("#FF0000", autoShape.LineColor);
        Assert.Equal(3.0, autoShape.LineWidth);
        Assert.Equal("Dash", autoShape.LineDashStyle);
    }

    [Fact]
    public void GetDetails_WithDefaultShape_ShouldNotFailOnLineFormat()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);

        var details = _provider.GetDetails(slide.Shapes[0], presentation);
        var autoShape = Assert.IsType<AutoShapeDetails>(details);

        Assert.NotNull(autoShape);
    }
}
