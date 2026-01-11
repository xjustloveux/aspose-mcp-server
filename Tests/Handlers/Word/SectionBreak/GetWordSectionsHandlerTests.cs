using Aspose.Words;
using AsposeMcpServer.Handlers.Word.SectionBreak;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.SectionBreak;

public class GetWordSectionsHandlerTests : WordHandlerTestBase
{
    private readonly GetWordSectionsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithMultipleSections(int sectionCount)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        for (var i = 0; i < sectionCount; i++)
        {
            builder.Write($"Section {i + 1} content");
            if (i < sectionCount - 1) builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        return doc;
    }

    #endregion

    #region Basic Get Sections Operations

    [Fact]
    public void Execute_ReturnsAllSections()
    {
        var doc = CreateDocumentWithMultipleSections(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"totalSections\": 3", result);
        Assert.Contains("sections", result);
    }

    [Fact]
    public void Execute_WithSectionIndex_ReturnsSpecificSection()
    {
        var doc = CreateDocumentWithMultipleSections(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("section", result);
        Assert.Contains("\"index\": 1", result);
    }

    [Fact]
    public void Execute_WithInvalidSectionIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithMultipleSections(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 10 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_ReturnsSectionBreakInfo()
    {
        var doc = CreateDocumentWithMultipleSections(2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("sectionBreak", result);
        Assert.Contains("type", result);
    }

    [Fact]
    public void Execute_ReturnsPageSetupInfo()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("pageSetup", result);
        Assert.Contains("paperSize", result);
        Assert.Contains("orientation", result);
        Assert.Contains("margins", result);
    }

    [Fact]
    public void Execute_ReturnsContentStatistics()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("contentStatistics", result);
        Assert.Contains("paragraphs", result);
        Assert.Contains("tables", result);
    }

    #endregion
}
