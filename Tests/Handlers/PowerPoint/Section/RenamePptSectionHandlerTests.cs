using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Section;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Section;

public class RenamePptSectionHandlerTests : PptHandlerTestBase
{
    private readonly RenamePptSectionHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Rename()
    {
        Assert.Equal("rename", _handler.Operation);
    }

    #endregion

    #region Various Section Indices

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_RenamesVariousSections(int sectionIndex)
    {
        var pres = CreatePresentationWithSections(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", sectionIndex },
            { "newName", $"Renamed {sectionIndex}" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal($"Renamed {sectionIndex}", pres.Sections[sectionIndex].Name);
        Assert.Equal(3, pres.Sections.Count);
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

    #region Basic Rename Operations

    [Fact]
    public void Execute_RenamesSection()
    {
        var pres = CreatePresentationWithSections(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 0 },
            { "newName", "Updated Section" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
        Assert.Equal("Updated Section", pres.Sections[0].Name);
        Assert.Equal(3, pres.Sections.Count);
    }

    [Fact]
    public void Execute_ReturnsSectionIndex()
    {
        var pres = CreatePresentationWithSections(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 1 },
            { "newName", "New Name" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal("New Name", pres.Sections[1].Name);
        Assert.Equal("Section 0", pres.Sections[0].Name);
        Assert.Equal("Section 2", pres.Sections[2].Name);
    }

    [Fact]
    public void Execute_ReturnsNewName()
    {
        var pres = CreatePresentationWithSections(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 0 },
            { "newName", "Introduction" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal("Introduction", pres.Sections[0].Name);
    }

    [Fact]
    public void Execute_UpdatesSectionName()
    {
        var pres = CreatePresentationWithSections(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 0 },
            { "newName", "Updated Name" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Updated Name", pres.Sections[0].Name);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutSectionIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSections(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "newName", "New Name" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutNewName_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSections(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidSectionIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSections(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 99 },
            { "newName", "New Name" }
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
            { "sectionIndex", -1 },
            { "newName", "New Name" }
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
            { "sectionIndex", 0 },
            { "newName", "New Name" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
