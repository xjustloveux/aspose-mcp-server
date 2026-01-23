using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Section;
using AsposeMcpServer.Results.PowerPoint.Section;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSectionsResult>(res);

        Assert.Equal(0, result.Count);
    }

    [Fact]
    public void Execute_WithNoSections_ReturnsMessage()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSectionsResult>(res);

        Assert.Contains("No sections found", result.Message);
    }

    [Fact]
    public void Execute_WithSections_ReturnsCount()
    {
        var pres = CreatePresentationWithSections(3);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSectionsResult>(res);

        Assert.Equal(3, result.Count);
    }

    [Fact]
    public void Execute_ReturnsResultType()
    {
        var pres = CreatePresentationWithSections(2);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSectionsResult>(res);

        Assert.NotNull(result);
        Assert.IsType<GetSectionsResult>(result);
    }

    #endregion

    #region Section Details

    [Fact]
    public void Execute_ReturnsSectionIndex()
    {
        var pres = CreatePresentationWithSections(2);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSectionsResult>(res);

        Assert.NotNull(result.Sections);
        Assert.Equal(0, result.Sections[0].Index);
        Assert.Equal(1, result.Sections[1].Index);
    }

    [Fact]
    public void Execute_ReturnsSectionName()
    {
        var pres = CreatePresentationWithSections(2);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSectionsResult>(res);

        Assert.NotNull(result.Sections);
        Assert.NotNull(result.Sections[0].Name);
    }

    [Fact]
    public void Execute_ReturnsStartSlideIndex()
    {
        var pres = CreatePresentationWithSections(2);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSectionsResult>(res);

        Assert.NotNull(result.Sections);
        Assert.True(result.Sections[0].StartSlideIndex >= 0);
    }

    [Fact]
    public void Execute_ReturnsSlideCount()
    {
        var pres = CreatePresentationWithSections(2);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSectionsResult>(res);

        Assert.NotNull(result.Sections);
        Assert.True(result.Sections[0].SlideCount >= 0);
    }

    #endregion
}
