using System.Text.Json.Nodes;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Section;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Section;

public class GetPptSectionsHandlerTests : PptHandlerTestBase
{
    private readonly GetPptSectionsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Presentation CreatePresentationWithSections(int sectionCount)
    {
        var pres = new Presentation();
        for (var i = 0; i < sectionCount * 2; i++)
            pres.Slides.AddEmptySlide(pres.LayoutSlides[0]);

        for (var i = 0; i < sectionCount; i++)
            pres.Sections.AddSection($"Section {i}", pres.Slides[i * 2]);

        return pres;
    }

    #endregion

    #region Get All Sections

    [Fact]
    public void Execute_WithNoSections_ReturnsEmptyResult()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"count\": 0", result);
    }

    [Fact]
    public void Execute_WithNoSections_ReturnsMessage()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("No sections found", result);
    }

    [Fact]
    public void Execute_WithSections_ReturnsCount()
    {
        var pres = CreatePresentationWithSections(3);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"count\": 3", result);
    }

    [Fact]
    public void Execute_ReturnsJsonFormat()
    {
        var pres = CreatePresentationWithSections(2);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonNode.Parse(result);
        Assert.NotNull(json);
    }

    #endregion

    #region Section Details

    [Fact]
    public void Execute_ReturnsSectionIndex()
    {
        var pres = CreatePresentationWithSections(2);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"index\": 0", result);
        Assert.Contains("\"index\": 1", result);
    }

    [Fact]
    public void Execute_ReturnsSectionName()
    {
        var pres = CreatePresentationWithSections(2);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("name", result);
    }

    [Fact]
    public void Execute_ReturnsStartSlideIndex()
    {
        var pres = CreatePresentationWithSections(2);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("startSlideIndex", result);
    }

    [Fact]
    public void Execute_ReturnsSlideCount()
    {
        var pres = CreatePresentationWithSections(2);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("slideCount", result);
    }

    #endregion
}
