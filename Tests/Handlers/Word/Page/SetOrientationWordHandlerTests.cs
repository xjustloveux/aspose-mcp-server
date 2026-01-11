using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Page;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Page;

public class SetOrientationWordHandlerTests : WordHandlerTestBase
{
    private readonly SetOrientationWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetOrientation()
    {
        Assert.Equal("set_orientation", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutOrientation_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithMultipleSections()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Section 1");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 2");
        return doc;
    }

    #endregion

    #region Basic Set Operations

    [Fact]
    public void Execute_SetsLandscapeOrientation()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "orientation", "Landscape" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("landscape", result.ToLower());
        Assert.Equal(Orientation.Landscape, doc.Sections[0].PageSetup.Orientation);
        AssertModified(context);
    }

    [Fact]
    public void Execute_SetsPortraitOrientation()
    {
        var doc = CreateDocumentWithText("Sample text.");
        doc.Sections[0].PageSetup.Orientation = Orientation.Landscape;
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "orientation", "Portrait" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("portrait", result.ToLower());
        Assert.Equal(Orientation.Portrait, doc.Sections[0].PageSetup.Orientation);
    }

    [Fact]
    public void Execute_WithSectionIndex_SetsOrientationOnSpecificSection()
    {
        var doc = CreateDocumentWithMultipleSections();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "orientation", "Landscape" },
            { "sectionIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("1 section", result.ToLower());
    }

    #endregion
}
