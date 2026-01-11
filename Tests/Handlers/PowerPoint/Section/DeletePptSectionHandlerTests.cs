using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Section;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Section;

public class DeletePptSectionHandlerTests : PptHandlerTestBase
{
    private readonly DeletePptSectionHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Delete()
    {
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

    [Fact]
    public void Execute_DeletesSection()
    {
        var pres = CreatePresentationWithSections(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 2 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("removed", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_DecreasesSectionCount()
    {
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

    [Fact]
    public void Execute_ReturnsSectionIndex()
    {
        var pres = CreatePresentationWithSections(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Section 1", result);
    }

    #endregion

    #region Keep Slides Parameter

    [Fact]
    public void Execute_WithKeepSlidesTrue_KeepsSlides()
    {
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

    [Fact]
    public void Execute_WithKeepSlidesFalse_RemovesSlidesWithSection()
    {
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

    [Fact]
    public void Execute_DefaultKeepSlides_KeepsSlides()
    {
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

    [Fact]
    public void Execute_ReturnsKeepSlidesStatus()
    {
        var pres = CreatePresentationWithSections(2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 1 },
            { "keepSlides", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("keep slides", result);
    }

    #endregion

    #region Various Section Indices

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_DeletesVariousSections(int sectionIndex)
    {
        var pres = CreatePresentationWithSections(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", sectionIndex }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("removed", result);
    }

    [Fact]
    public void Execute_DeletesOnlySection()
    {
        var pres = CreatePresentationWithSections(1);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("removed", result);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutSectionIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSections(3);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidSectionIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSections(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNegativeSectionIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSections(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", -1 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNoSections_ThrowsArgumentException()
    {
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
