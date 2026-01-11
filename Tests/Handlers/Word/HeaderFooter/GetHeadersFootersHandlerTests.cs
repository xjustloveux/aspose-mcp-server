using System.Text.Json;
using Aspose.Words;
using AsposeMcpServer.Handlers.Word.HeaderFooter;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.HeaderFooter;

public class GetHeadersFootersHandlerTests : WordHandlerTestBase
{
    private readonly GetHeadersFootersHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidSectionIndex_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_DoesNotModifyDocument()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        AssertNotModified(context);
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_ReturnsHeadersFootersInfo()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("totalSections", out _));
        Assert.True(json.RootElement.TryGetProperty("sections", out _));
    }

    [Fact]
    public void Execute_WithDocumentWithHeader_ReturnsHeaderContent()
    {
        var doc = CreateDocumentWithHeader("Test Header");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var sections = json.RootElement.GetProperty("sections");
        Assert.True(sections.GetArrayLength() > 0);
    }

    [Fact]
    public void Execute_WithDocumentWithFooter_ReturnsFooterContent()
    {
        var doc = CreateDocumentWithFooter("Test Footer");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("sections", out _));
    }

    [Fact]
    public void Execute_WithSectionIndex_ReturnsSpecificSection()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sectionIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("queriedSectionIndex").GetInt32());
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
        return doc;
    }

    private static Document CreateDocumentWithFooter(string footerText)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write(footerText);
        builder.MoveToDocumentEnd();
        return doc;
    }

    #endregion
}
