using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Handlers.Pdf.Info;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Info;

public class GetPdfContentHandlerTests : PdfHandlerTestBase
{
    private readonly GetPdfContentHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetContent()
    {
        Assert.Equal("get_content", _handler.Operation);
    }

    #endregion

    #region Basic Get Content Operations

    [SkippableFact]
    public void Execute_ReturnsContent()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var doc = CreatePdfWithText("Test content for extraction");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("content", result);
        Assert.Contains("totalPages", result);
    }

    [SkippableFact]
    public void Execute_ReturnsJsonFormat()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var doc = CreatePdfWithText("Sample text");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonNode.Parse(result);
        Assert.NotNull(json);
        Assert.NotNull(json["totalPages"]);
        Assert.NotNull(json["content"]);
    }

    [SkippableFact]
    public void Execute_WithMultiplePages_ReturnsAllContent()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var doc = CreateDocumentWithPages(3);
        AddTextToPage(doc, 1, "Page 1 content");
        AddTextToPage(doc, 2, "Page 2 content");
        AddTextToPage(doc, 3, "Page 3 content");

        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonNode.Parse(result);
        Assert.NotNull(json);
        Assert.Equal(3, json["totalPages"]?.GetValue<int>());
        Assert.Equal(3, json["extractedPages"]?.GetValue<int>());
    }

    #endregion

    #region Page Index Parameter

    [SkippableFact]
    public void Execute_WithPageIndex_ReturnsSpecificPageContent()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var doc = CreateDocumentWithPages(3);
        AddTextToPage(doc, 2, "Page 2 specific content");

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 2 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonNode.Parse(result);
        Assert.NotNull(json);
        Assert.Equal(2, json["pageIndex"]?.GetValue<int>());
    }

    [Fact]
    public void Execute_WithInvalidPageIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("pageIndex must be between", ex.Message);
    }

    [Fact]
    public void Execute_WithPageIndexZero_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("pageIndex must be between", ex.Message);
    }

    #endregion

    #region MaxPages Parameter

    [SkippableFact]
    public void Execute_WithMaxPages_LimitsExtraction()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var doc = CreateDocumentWithPages(10);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "maxPages", 3 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonNode.Parse(result);
        Assert.NotNull(json);
        Assert.Equal(10, json["totalPages"]?.GetValue<int>());
        Assert.Equal(3, json["extractedPages"]?.GetValue<int>());
        Assert.True(json["truncated"]?.GetValue<bool>());
    }

    [SkippableFact]
    public void Execute_WithMaxPagesGreaterThanTotal_ReturnsAllPages()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "maxPages", 100 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonNode.Parse(result);
        Assert.NotNull(json);
        Assert.Equal(3, json["extractedPages"]?.GetValue<int>());
        Assert.False(json["truncated"]?.GetValue<bool>());
    }

    #endregion

    #region Helper Methods

    private static Document CreatePdfWithText(string text)
    {
        var doc = new Document();
        var page = doc.Pages.Add();
        page.Paragraphs.Add(new TextFragment(text));
        return doc;
    }

    private static void AddTextToPage(Document doc, int pageIndex, string text)
    {
        var page = doc.Pages[pageIndex];
        page.Paragraphs.Add(new TextFragment(text));
    }

    #endregion
}
