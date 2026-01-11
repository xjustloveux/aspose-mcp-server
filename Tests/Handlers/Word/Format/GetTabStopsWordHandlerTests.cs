using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Format;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Format;

public class GetTabStopsWordHandlerTests : WordHandlerTestBase
{
    private readonly GetTabStopsWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetTabStops()
    {
        Assert.Equal("get_tab_stops", _handler.Operation);
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
            { "sectionIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithTabStops()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Column 1\tColumn 2");
        builder.CurrentParagraph.ParagraphFormat.TabStops.Add(new TabStop(144.0, TabAlignment.Left, TabLeader.None));
        return doc;
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_ReturnsTabStops()
    {
        var doc = CreateDocumentWithTabStops();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("tabStops", result);
        Assert.Contains("count", result);
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_ReturnsJsonFormat()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("{", result);
        Assert.Contains("}", result);
        Assert.Contains("location", result);
    }

    [Fact]
    public void Execute_WithAllParagraphs_ReturnsAllTabStops()
    {
        var doc = CreateDocumentWithTabStops();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "allParagraphs", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("paragraphCount", result);
    }

    #endregion
}
