using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Content;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Content;

public class GetWordContentDetailedHandlerTests : WordHandlerTestBase
{
    private readonly GetWordContentDetailedHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetContentDetailed()
    {
        Assert.Equal("get_content_detailed", _handler.Operation);
    }

    #endregion

    #region Combined Parameters

    [Fact]
    public void Execute_WithBothHeadersAndFooters_ShowsBoth()
    {
        var doc = CreateDocumentWithHeaderAndFooter("Header", "Footer");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "includeHeaders", true },
            { "includeFooters", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Headers", result);
        Assert.Contains("Footers", result);
    }

    #endregion

    #region Multiple Sections

    [Fact]
    public void Execute_WithMultipleSections_ReturnsAllContent()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Section 1 content");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 2 content");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Section 1 content", result);
        Assert.Contains("Section 2 content", result);
    }

    #endregion

    #region Basic Content Retrieval

    [Fact]
    public void Execute_ReturnsDetailedContent()
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Detailed Document Content", result);
    }

    [Fact]
    public void Execute_ReturnsBodyContent()
    {
        var doc = CreateDocumentWithText("Body content here");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Body Content", result);
        Assert.Contains("Body content here", result);
    }

    [Fact]
    public void Execute_DoesNotMarkAsModified()
    {
        var doc = CreateDocumentWithText("Read only");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        Assert.False(context.IsModified);
    }

    #endregion

    #region Include Headers Parameter

    [Fact]
    public void Execute_WithIncludeHeadersTrue_ShowsHeadersSection()
    {
        var doc = CreateDocumentWithHeader("Header Text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "includeHeaders", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Headers", result);
    }

    [Fact]
    public void Execute_WithIncludeHeadersFalse_HidesHeaders()
    {
        var doc = CreateDocumentWithHeader("Header Text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "includeHeaders", false }
        });

        var result = _handler.Execute(context, parameters);

        Assert.DoesNotContain("--- Headers ---", result);
    }

    [Fact]
    public void Execute_DefaultIncludeHeaders_DoesNotShowHeaders()
    {
        var doc = CreateDocumentWithHeader("Header Text");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.DoesNotContain("--- Headers ---", result);
    }

    #endregion

    #region Include Footers Parameter

    [Fact]
    public void Execute_WithIncludeFootersTrue_ShowsFootersSection()
    {
        var doc = CreateDocumentWithFooter("Footer Text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "includeFooters", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Footers", result);
    }

    [Fact]
    public void Execute_WithIncludeFootersFalse_HidesFooters()
    {
        var doc = CreateDocumentWithFooter("Footer Text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "includeFooters", false }
        });

        var result = _handler.Execute(context, parameters);

        Assert.DoesNotContain("--- Footers ---", result);
    }

    [Fact]
    public void Execute_DefaultIncludeFooters_DoesNotShowFooters()
    {
        var doc = CreateDocumentWithFooter("Footer Text");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.DoesNotContain("--- Footers ---", result);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithHeader(string headerText)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write(headerText);
        builder.MoveToDocumentEnd();
        builder.Writeln("Body content");
        return doc;
    }

    private static Document CreateDocumentWithFooter(string footerText)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write(footerText);
        builder.MoveToDocumentEnd();
        builder.Writeln("Body content");
        return doc;
    }

    private static Document CreateDocumentWithHeaderAndFooter(string headerText, string footerText)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write(headerText);
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write(footerText);
        builder.MoveToDocumentEnd();
        builder.Writeln("Body content");
        return doc;
    }

    #endregion
}
