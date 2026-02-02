using AsposeMcpServer.Handlers.PowerPoint.Section;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Section;

public class AddPptSectionHandlerTests : PptHandlerTestBase
{
    private readonly AddPptSectionHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Various Slide Indices

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_AddsSectionAtVariousSlideIndices(int slideIndex)
    {
        var pres = CreatePresentationWithSlides(3);
        var initialCount = pres.Sections.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", $"Section at {slideIndex}" },
            { "slideIndex", slideIndex }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(pres.Sections.Count > initialCount);
        var found = false;
        foreach (var section in pres.Sections)
        {
            if (section.Name != $"Section at {slideIndex}") continue;
            found = true;
            break;
        }

        Assert.True(found);
    }

    #endregion

    #region Multiple Sections

    [Fact]
    public void Execute_AddsMultipleSections()
    {
        var pres = CreatePresentationWithSlides(5);
        var context = CreateContext(pres);

        _handler.Execute(context, CreateParameters(new Dictionary<string, object?>
        {
            { "name", "Section 1" },
            { "slideIndex", 0 }
        }));

        _handler.Execute(context, CreateParameters(new Dictionary<string, object?>
        {
            { "name", "Section 2" },
            { "slideIndex", 2 }
        }));

        Assert.Equal(2, pres.Sections.Count);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsSection()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "New Section" },
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
        var section = Assert.Single(pres.Sections);
        Assert.Equal("New Section", section.Name);
    }

    [Fact]
    public void Execute_ReturnsSectionName()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "Introduction" },
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        var section = Assert.Single(pres.Sections);
        Assert.Equal("Introduction", section.Name);
    }

    [Fact]
    public void Execute_ReturnsSlideIndex()
    {
        var pres = CreatePresentationWithSlides(3);
        var initialCount = pres.Sections.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "Section" },
            { "slideIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(pres.Sections.Count > initialCount);
        foreach (var section in pres.Sections)
        {
            if (section.Name != "Section") continue;
            Assert.Equal(pres.Slides[1], section.StartedFromSlide);
            return;
        }

        Assert.Fail("Section not found");
    }

    [Fact]
    public void Execute_IncreasesSectionCount()
    {
        var pres = CreatePresentationWithSlides(3);
        var initialCount = pres.Sections.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "New Section" },
            { "slideIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialCount + 1, pres.Sections.Count);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutName_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutSlideIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "Section" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "Section" },
            { "slideIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNegativeSlideIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "Section" },
            { "slideIndex", -1 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
