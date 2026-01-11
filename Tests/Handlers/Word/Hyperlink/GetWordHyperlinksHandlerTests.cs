using System.Text.Json;
using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Hyperlink;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(0, json.RootElement.GetProperty("count").GetInt32());
        Assert.Contains("No hyperlinks found", json.RootElement.GetProperty("message").GetString());
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_GetsHyperlinks()
    {
        var doc = CreateDocumentWithHyperlinks(2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.True(json.RootElement.TryGetProperty("count", out _));
    }

    [Fact]
    public void Execute_ReturnsCorrectCount()
    {
        var doc = CreateDocumentWithHyperlinks(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(3, json.RootElement.GetProperty("count").GetInt32());
    }

    [Fact]
    public void Execute_ReturnsHyperlinksArray()
    {
        var doc = CreateDocumentWithHyperlinks(2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(2, json.RootElement.GetProperty("hyperlinks").GetArrayLength());
    }

    #endregion

    #region Hyperlink Details

    [Fact]
    public void Execute_ReturnsHyperlinkIndex()
    {
        var doc = CreateDocumentWithHyperlinks(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);
        var firstHyperlink = json.RootElement.GetProperty("hyperlinks")[0];

        Assert.Equal(0, firstHyperlink.GetProperty("index").GetInt32());
    }

    [Fact]
    public void Execute_ReturnsDisplayText()
    {
        var doc = CreateDocumentWithHyperlink("Click Me", "https://example.com");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);
        var firstHyperlink = json.RootElement.GetProperty("hyperlinks")[0];

        Assert.Equal("Click Me", firstHyperlink.GetProperty("displayText").GetString());
    }

    [Fact]
    public void Execute_ReturnsAddress()
    {
        var doc = CreateDocumentWithHyperlink("Link", "https://test.example.com");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);
        var firstHyperlink = json.RootElement.GetProperty("hyperlinks")[0];

        Assert.Equal("https://test.example.com", firstHyperlink.GetProperty("address").GetString());
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
