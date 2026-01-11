using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Section;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("renamed", result);
        Assert.Equal($"Renamed {sectionIndex}", pres.Sections[sectionIndex].Name);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("renamed", result);
        AssertModified(context);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Section 1", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Introduction", result);
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
