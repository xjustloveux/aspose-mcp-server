using System.Runtime.Versioning;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Section;
using AsposeMcpServer.Results.PowerPoint.Section;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Section;

[SupportedOSPlatform("windows")]
public class GetPptSectionsHandlerTests : PptHandlerTestBase
{
    private readonly GetPptSectionsHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_Get()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_WithNoSections_ReturnsEmptyResult()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSectionsResult>(res);

        Assert.Equal(0, result.Count);
    }

    [SkippableFact]
    public void Execute_WithNoSections_ReturnsMessage()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSectionsResult>(res);

        Assert.Contains("No sections found", result.Message);
    }

    [SkippableFact]
    public void Execute_WithSections_ReturnsCount()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSections(3);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSectionsResult>(res);

        Assert.Equal(3, result.Count);
    }

    [SkippableFact]
    public void Execute_ReturnsResultType()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_ReturnsSectionIndex()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSections(2);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSectionsResult>(res);

        Assert.NotNull(result.Sections);
        Assert.Equal(0, result.Sections[0].Index);
        Assert.Equal(1, result.Sections[1].Index);
    }

    [SkippableFact]
    public void Execute_ReturnsSectionName()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSections(2);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSectionsResult>(res);

        Assert.NotNull(result.Sections);
        Assert.NotNull(result.Sections[0].Name);
    }

    [SkippableFact]
    public void Execute_ReturnsStartSlideIndex()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSections(2);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSectionsResult>(res);

        Assert.NotNull(result.Sections);
        Assert.True(result.Sections[0].StartSlideIndex >= 0);
    }

    [SkippableFact]
    public void Execute_ReturnsSlideCount()
    {
        SkipIfNotWindows();
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
