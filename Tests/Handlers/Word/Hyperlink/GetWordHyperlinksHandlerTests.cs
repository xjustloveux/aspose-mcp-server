using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Hyperlink;
using AsposeMcpServer.Results.Word.Hyperlink;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Hyperlink;

public class GetWordHyperlinksHandlerTests : WordHandlerTestBase
{
    private readonly GetWordHyperlinksHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region No Hyperlinks

    [Fact]
    public void Execute_NoHyperlinks_ReturnsEmptyResult()
    {
        var doc = CreateDocumentWithText("No hyperlinks here");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetHyperlinksResult>(res);

        Assert.Equal(0, result.Count);
        Assert.NotNull(result.Message);
        Assert.Contains("No hyperlinks found", result.Message);
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_GetsHyperlinks()
    {
        var doc = CreateDocumentWithHyperlinks(2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetHyperlinksResult>(res);

        Assert.True(result.Count >= 0);
    }

    [Fact]
    public void Execute_ReturnsCorrectCount()
    {
        var doc = CreateDocumentWithHyperlinks(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetHyperlinksResult>(res);

        Assert.Equal(3, result.Count);
    }

    [Fact]
    public void Execute_ReturnsHyperlinksArray()
    {
        var doc = CreateDocumentWithHyperlinks(2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetHyperlinksResult>(res);

        Assert.Equal(2, result.Hyperlinks.Count);
    }

    #endregion

    #region Hyperlink Details

    [Fact]
    public void Execute_ReturnsHyperlinkIndex()
    {
        var doc = CreateDocumentWithHyperlinks(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetHyperlinksResult>(res);
        var firstHyperlink = result.Hyperlinks[0];

        Assert.Equal(0, firstHyperlink.Index);
    }

    [Fact]
    public void Execute_ReturnsDisplayText()
    {
        var doc = CreateDocumentWithHyperlink("Click Me", "https://example.com");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetHyperlinksResult>(res);
        var firstHyperlink = result.Hyperlinks[0];

        Assert.Equal("Click Me", firstHyperlink.DisplayText);
    }

    [Fact]
    public void Execute_ReturnsAddress()
    {
        var doc = CreateDocumentWithHyperlink("Link", "https://test.example.com");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetHyperlinksResult>(res);
        var firstHyperlink = result.Hyperlinks[0];

        Assert.Equal("https://test.example.com", firstHyperlink.Address);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithHyperlinks(int count)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        for (var i = 0; i < count; i++)
        {
            builder.InsertHyperlink($"Link {i}", $"https://example{i}.com", false);
            builder.InsertParagraph();
        }

        return doc;
    }

    private static Document CreateDocumentWithHyperlink(string text, string url)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertHyperlink(text, url, false);
        return doc;
    }

    #endregion
}
