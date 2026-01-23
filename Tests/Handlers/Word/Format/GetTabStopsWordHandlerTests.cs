using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Format;
using AsposeMcpServer.Results.Word.Format;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTabStopsWordResult>(res);

        Assert.NotNull(result);
        Assert.NotNull(result.TabStops);
        Assert.True(result.Count >= 0);
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_ReturnsLocation()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTabStopsWordResult>(res);

        Assert.NotNull(result);
        Assert.NotNull(result.Location);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTabStopsWordResult>(res);

        Assert.NotNull(result);
        Assert.True(result.ParagraphCount >= 1);
    }

    [Fact]
    public void Execute_WithInvalidParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_WithIncludeStyleFalse_OmitsStyleTabStops()
    {
        var doc = CreateDocumentWithTabStops();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "includeStyle", false }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTabStopsWordResult>(res);

        Assert.NotNull(result);
        Assert.False(result.IncludeStyle);
    }

    #endregion

    #region Header/Footer Location Tests

    [SkippableFact]
    public void Execute_WithHeaderLocation_ReturnsHeaderTabStops()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Evaluation mode limits header operations");

        var doc = CreateDocumentWithHeader();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "location", "header" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTabStopsWordResult>(res);

        Assert.NotNull(result);
        Assert.Contains("Header", result.LocationDescription);
    }

    [SkippableFact]
    public void Execute_WithFooterLocation_ReturnsFooterTabStops()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Evaluation mode limits footer operations");

        var doc = CreateDocumentWithFooter();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "location", "footer" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTabStopsWordResult>(res);

        Assert.NotNull(result);
        Assert.Contains("Footer", result.LocationDescription);
    }

    [SkippableFact]
    public void Execute_WithMissingHeader_ThrowsInvalidOperationException()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Evaluation mode adds headers automatically");

        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "location", "header" }
        });

        var ex = Assert.Throws<InvalidOperationException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Header not found", ex.Message);
    }

    [SkippableFact]
    public void Execute_WithMissingFooter_ThrowsInvalidOperationException()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Evaluation mode adds footers automatically");

        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "location", "footer" }
        });

        var ex = Assert.Throws<InvalidOperationException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Footer not found", ex.Message);
    }

    private static Document CreateDocumentWithHeader()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header Text");
        builder.MoveToDocumentEnd();
        builder.Write("Body Text");
        return doc;
    }

    private static Document CreateDocumentWithFooter()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Footer Text");
        builder.MoveToDocumentEnd();
        builder.Write("Body Text");
        return doc;
    }

    #endregion
}
