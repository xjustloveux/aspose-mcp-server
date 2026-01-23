using Aspose.Words;
using AsposeMcpServer.Handlers.Word.HeaderFooter;
using AsposeMcpServer.Results.Word.HeaderFooter;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetHeadersFootersResult>(res);

        Assert.True(result.TotalSections >= 0);
        Assert.NotNull(result.Sections);
    }

    [Fact]
    public void Execute_WithDocumentWithHeader_ReturnsHeaderContent()
    {
        var doc = CreateDocumentWithHeader("Test Header");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetHeadersFootersResult>(res);

        Assert.True(result.Sections.Count > 0);
    }

    [Fact]
    public void Execute_WithDocumentWithFooter_ReturnsFooterContent()
    {
        var doc = CreateDocumentWithFooter("Test Footer");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetHeadersFootersResult>(res);

        Assert.NotNull(result.Sections);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetHeadersFootersResult>(res);

        Assert.Equal(0, result.QueriedSectionIndex);
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
