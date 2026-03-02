using System.Runtime.Versioning;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Section;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Section;

[SupportedOSPlatform("windows")]
public class DeletePptSectionHandlerTests : PptHandlerTestBase
{
    private readonly DeletePptSectionHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_Delete()
    {
        SkipIfNotWindows();
        Assert.Equal("delete", _handler.Operation);
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

    #region Basic Delete Operations

    [SkippableFact]
    public void Execute_DeletesSection()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSections(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 2 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("removed", result.Message);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_DecreasesSectionCount()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSections(3);
        var initialCount = pres.Sections.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 2 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialCount - 1, pres.Sections.Count);
    }

    [SkippableFact]
    public void Execute_ReturnsSectionIndex()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSections(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Section 1", result.Message);
    }

    #endregion

    #region Keep Slides Parameter

    [SkippableFact]
    public void Execute_WithKeepSlidesTrue_KeepsSlides()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSections(2);
        var initialSlideCount = pres.Slides.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 1 },
            { "keepSlides", true }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialSlideCount, pres.Slides.Count);
    }

    [SkippableFact]
    public void Execute_WithKeepSlidesFalse_RemovesSlidesWithSection()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSections(2);
        var initialSlideCount = pres.Slides.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 1 },
            { "keepSlides", false }
        });

        _handler.Execute(context, parameters);

        Assert.True(pres.Slides.Count < initialSlideCount);
    }

    [SkippableFact]
    public void Execute_DefaultKeepSlides_KeepsSlides()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSections(2);
        var initialSlideCount = pres.Slides.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 1 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialSlideCount, pres.Slides.Count);
    }

    [SkippableFact]
    public void Execute_ReturnsKeepSlidesStatus()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSections(2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 1 },
            { "keepSlides", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("keep slides", result.Message);
    }

    #endregion

    #region Various Section Indices

    [SkippableTheory]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_DeletesVariousSections(int sectionIndex)
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSections(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", sectionIndex }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("removed", result.Message);
    }

    [SkippableFact]
    public void Execute_DeletesOnlySection()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSections(1);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("removed", result.Message);
    }

    #endregion

    #region Error Handling

    [SkippableFact]
    public void Execute_WithoutSectionIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSections(3);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithInvalidSectionIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSections(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithNegativeSectionIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSections(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", -1 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithNoSections_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
