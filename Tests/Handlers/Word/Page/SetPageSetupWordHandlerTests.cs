using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Page;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Page;

public class SetPageSetupWordHandlerTests : WordHandlerTestBase
{
    private readonly SetPageSetupWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetPageSetup()
    {
        Assert.Equal("set_page_setup", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidSectionIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 99 },
            { "top", 72.0 }
        });

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
    public void Execute_SetsMultiplePageSetupOptions()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "top", 72.0 },
            { "bottom", 72.0 },
            { "orientation", "Landscape" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("page setup updated", result.ToLower());
        Assert.Equal(72.0, doc.Sections[0].PageSetup.TopMargin);
        Assert.Equal(72.0, doc.Sections[0].PageSetup.BottomMargin);
        Assert.Equal(Orientation.Landscape, doc.Sections[0].PageSetup.Orientation);
        AssertModified(context);
    }

    [Fact]
    public void Execute_SetsAllMargins()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "top", 50.0 },
            { "bottom", 50.0 },
            { "left", 40.0 },
            { "right", 40.0 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(50.0, doc.Sections[0].PageSetup.TopMargin);
        Assert.Equal(50.0, doc.Sections[0].PageSetup.BottomMargin);
        Assert.Equal(40.0, doc.Sections[0].PageSetup.LeftMargin);
        Assert.Equal(40.0, doc.Sections[0].PageSetup.RightMargin);
    }

    [Fact]
    public void Execute_WithSectionIndex_SetsOnSpecificSection()
    {
        var doc = CreateDocumentWithMultipleSections();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 1 },
            { "top", 100.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("top margin: 100", result.ToLower());
    }

    #endregion
}
