using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Hyperlink;
using PptHyperlink = Aspose.Slides.Hyperlink;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Hyperlink;

public class PptHyperlinkHelperTests
{
    #region GetHyperlinksFromSlide Tests

    [Fact]
    public void GetHyperlinksFromSlide_WithNoHyperlinks_ReturnsEmptyList()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];

        var result = PptHyperlinkHelper.GetHyperlinksFromSlide(presentation, slide);

        Assert.Empty(result);
    }

    [Fact]
    public void GetHyperlinksFromSlide_WithShapeHyperlink_ReturnsHyperlink()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 50);
        shape.HyperlinkClick = new PptHyperlink("https://example.com");

        var result = PptHyperlinkHelper.GetHyperlinksFromSlide(presentation, slide);

        Assert.Single(result);
    }

    [Fact]
    public void GetHyperlinksFromSlide_WithMouseOverHyperlink_ReturnsHyperlink()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 50);
        shape.HyperlinkMouseOver = new PptHyperlink("https://mouseover.com");

        var result = PptHyperlinkHelper.GetHyperlinksFromSlide(presentation, slide);

        Assert.Single(result);
    }

    [Fact]
    public void GetHyperlinksFromSlide_WithTextHyperlink_ReturnsHyperlink()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 50);
        var portion = shape.TextFrame.Paragraphs[0].Portions[0];
        portion.Text = "Click here";
        portion.PortionFormat.HyperlinkClick = new PptHyperlink("https://text-link.com");

        var result = PptHyperlinkHelper.GetHyperlinksFromSlide(presentation, slide);

        Assert.Single(result);
    }

    [Fact]
    public void GetHyperlinksFromSlide_WithInternalSlideLink_ReturnsSlideReference()
    {
        using var presentation = new Presentation();
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 50);
        shape.HyperlinkClick = new PptHyperlink(presentation.Slides[1]);

        var result = PptHyperlinkHelper.GetHyperlinksFromSlide(presentation, slide);

        Assert.Single(result);
    }

    #endregion

    #region CreateHyperlink Tests

    [Fact]
    public void CreateHyperlink_WithUrl_ReturnsExternalHyperlink()
    {
        using var presentation = new Presentation();

        var (hyperlink, description) = PptHyperlinkHelper.CreateHyperlink(presentation, "https://example.com", null);

        Assert.NotNull(hyperlink);
        Assert.Equal("https://example.com", description);
    }

    [Fact]
    public void CreateHyperlink_WithValidSlideIndex_ReturnsInternalHyperlink()
    {
        using var presentation = new Presentation();
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);

        var (hyperlink, description) = PptHyperlinkHelper.CreateHyperlink(presentation, null, 1);

        Assert.NotNull(hyperlink);
        Assert.Equal("Slide 1", description);
    }

    [Fact]
    public void CreateHyperlink_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        using var presentation = new Presentation();

        var ex = Assert.Throws<ArgumentException>(() =>
            PptHyperlinkHelper.CreateHyperlink(presentation, null, 10));

        Assert.Contains("slideTargetIndex must be between", ex.Message);
    }

    [Fact]
    public void CreateHyperlink_WithNegativeSlideIndex_ThrowsArgumentException()
    {
        using var presentation = new Presentation();

        var ex = Assert.Throws<ArgumentException>(() =>
            PptHyperlinkHelper.CreateHyperlink(presentation, null, -1));

        Assert.Contains("slideTargetIndex must be between", ex.Message);
    }

    [Fact]
    public void CreateHyperlink_WithNoUrlAndNoSlideIndex_ThrowsArgumentException()
    {
        using var presentation = new Presentation();

        var ex = Assert.Throws<ArgumentException>(() =>
            PptHyperlinkHelper.CreateHyperlink(presentation, null, null));

        Assert.Contains("Either url or slideTargetIndex must be provided", ex.Message);
    }

    [Fact]
    public void CreateHyperlink_WithEmptyUrl_ThrowsArgumentException()
    {
        using var presentation = new Presentation();

        var ex = Assert.Throws<ArgumentException>(() =>
            PptHyperlinkHelper.CreateHyperlink(presentation, "", null));

        Assert.Contains("Either url or slideTargetIndex must be provided", ex.Message);
    }

    #endregion
}
