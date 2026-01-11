using System.Text.Json;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Slide;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Slide;

public class GetPptSlidesInfoHandlerTests : PptHandlerTestBase
{
    private readonly GetPptSlidesInfoHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetInfo()
    {
        Assert.Equal("get_info", _handler.Operation);
    }

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_DoesNotModifyPresentation()
    {
        var pres = CreatePresentationWithSlides(3);
        var initialCount = pres.Slides.Count;
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        Assert.Equal(initialCount, pres.Slides.Count);
        AssertNotModified(context);
    }

    #endregion

    #region Basic Info Retrieval

    [Fact]
    public void Execute_ReturnsSlideInfo()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("count", out var count));
        Assert.Equal(3, count.GetInt32());
        AssertNotModified(context);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(3)]
    [InlineData(5)]
    [InlineData(10)]
    public void Execute_ReturnsCorrectSlideCount(int slideCount)
    {
        var pres = CreatePresentationWithSlides(slideCount);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(slideCount, json.RootElement.GetProperty("count").GetInt32());
        AssertNotModified(context);
    }

    #endregion

    #region Slides Array

    [Fact]
    public void Execute_ReturnsSlidesArray()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("slides", out var slides));
        Assert.Equal(JsonValueKind.Array, slides.ValueKind);
        Assert.Equal(3, slides.GetArrayLength());
    }

    [Fact]
    public void Execute_SlidesContainIndex()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var slides = json.RootElement.GetProperty("slides");
        var firstSlide = slides[0];
        Assert.True(firstSlide.TryGetProperty("index", out var index));
        Assert.Equal(0, index.GetInt32());
    }

    [Fact]
    public void Execute_SlidesContainLayoutInfo()
    {
        var pres = CreatePresentationWithSlides(1);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var slides = json.RootElement.GetProperty("slides");
        var firstSlide = slides[0];
        Assert.True(firstSlide.TryGetProperty("layoutType", out _));
        Assert.True(firstSlide.TryGetProperty("layoutName", out _));
    }

    [Fact]
    public void Execute_SlidesContainShapesCount()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var slides = json.RootElement.GetProperty("slides");
        var firstSlide = slides[0];
        Assert.True(firstSlide.TryGetProperty("shapesCount", out var shapesCount));
        Assert.True(shapesCount.GetInt32() >= 1);
    }

    [Fact]
    public void Execute_SlidesContainHiddenProperty()
    {
        var pres = CreatePresentationWithSlides(1);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var slides = json.RootElement.GetProperty("slides");
        var firstSlide = slides[0];
        Assert.True(firstSlide.TryGetProperty("hidden", out _));
    }

    #endregion

    #region Available Layouts

    [Fact]
    public void Execute_ReturnsAvailableLayouts()
    {
        var pres = CreatePresentationWithSlides(1);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("availableLayouts", out var layouts));
        Assert.Equal(JsonValueKind.Array, layouts.ValueKind);
        Assert.True(layouts.GetArrayLength() > 0);
    }

    [Fact]
    public void Execute_LayoutsContainNameAndType()
    {
        var pres = CreatePresentationWithSlides(1);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var layouts = json.RootElement.GetProperty("availableLayouts");
        var firstLayout = layouts[0];
        Assert.True(firstLayout.TryGetProperty("index", out _));
        Assert.True(firstLayout.TryGetProperty("name", out _));
        Assert.True(firstLayout.TryGetProperty("type", out _));
    }

    #endregion
}
