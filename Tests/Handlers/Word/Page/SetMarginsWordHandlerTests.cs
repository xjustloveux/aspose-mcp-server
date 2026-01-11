using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Page;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Page;

public class SetMarginsWordHandlerTests : WordHandlerTestBase
{
    private readonly SetMarginsWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetMargins()
    {
        Assert.Equal("set_margins", _handler.Operation);
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
    public void Execute_SetsAllMargins()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "top", 72.0 },
            { "bottom", 72.0 },
            { "left", 54.0 },
            { "right", 54.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("margins updated", result.ToLower());
        Assert.Equal(72.0, doc.Sections[0].PageSetup.TopMargin);
        Assert.Equal(72.0, doc.Sections[0].PageSetup.BottomMargin);
        Assert.Equal(54.0, doc.Sections[0].PageSetup.LeftMargin);
        Assert.Equal(54.0, doc.Sections[0].PageSetup.RightMargin);
        AssertModified(context);
    }

    [Fact]
    public void Execute_SetsTopMarginOnly()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "top", 100.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("margins updated", result.ToLower());
        Assert.Equal(100.0, doc.Sections[0].PageSetup.TopMargin);
    }

    [Fact]
    public void Execute_WithSectionIndex_SetsMarginsOnSpecificSection()
    {
        var doc = CreateDocumentWithMultipleSections();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 1 },
            { "top", 50.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("1 section", result.ToLower());
    }

    #endregion

    #region Boundary Condition Tests

    [Fact]
    public void Execute_WithZeroMargins_SetsZeroMargins()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "top", 0.0 },
            { "bottom", 0.0 },
            { "left", 0.0 },
            { "right", 0.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("margins updated", result.ToLower());
        Assert.Equal(0.0, doc.Sections[0].PageSetup.TopMargin);
    }

    [Fact]
    public void Execute_WithNegativeMargin_AcceptsValue()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "top", -10.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("margins updated", result.ToLower());
    }

    [Fact]
    public void Execute_WithNegativeSectionIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithMultipleSections();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", -1 },
            { "top", 50.0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("sectionIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidSectionIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithMultipleSections();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 99 },
            { "top", 50.0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("sectionIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithoutAnyMarginParameter_ReturnsNoChangesMessage()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("margins updated", result.ToLower());
    }

    #endregion
}
