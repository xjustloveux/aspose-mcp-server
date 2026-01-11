using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Page;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Page;

public class SetPageNumberWordHandlerTests : WordHandlerTestBase
{
    private readonly SetPageNumberWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetPageNumber()
    {
        Assert.Equal("set_page_number", _handler.Operation);
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
            { "pageNumberFormat", "arabic" },
            { "sectionIndex", 99 }
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
    public void Execute_SetsArabicPageNumberFormat()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageNumberFormat", "arabic" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("page number settings updated", result.ToLower());
        Assert.Equal(NumberStyle.Arabic, doc.Sections[0].PageSetup.PageNumberStyle);
        AssertModified(context);
    }

    [Fact]
    public void Execute_SetsRomanPageNumberFormat()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageNumberFormat", "roman" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(NumberStyle.UppercaseRoman, doc.Sections[0].PageSetup.PageNumberStyle);
    }

    [Fact]
    public void Execute_SetsStartingPageNumber()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startingPageNumber", 5 }
        });

        _handler.Execute(context, parameters);

        Assert.True(doc.Sections[0].PageSetup.RestartPageNumbering);
        Assert.Equal(5, doc.Sections[0].PageSetup.PageStartingNumber);
    }

    [Fact]
    public void Execute_WithSectionIndex_SetsOnSpecificSection()
    {
        var doc = CreateDocumentWithMultipleSections();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageNumberFormat", "letter" },
            { "sectionIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("1 section", result.ToLower());
    }

    #endregion
}
