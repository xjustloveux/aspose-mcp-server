using System.Text.Json;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Handlers.Pdf.Bookmark;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Bookmark;

public class GetPdfBookmarksHandlerTests : PdfHandlerTestBase
{
    private readonly GetPdfBookmarksHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region No Bookmarks

    [Fact]
    public void Execute_NoBookmarks_ReturnsEmptyResult()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(0, json.RootElement.GetProperty("count").GetInt32());
        Assert.Contains("No bookmarks found", json.RootElement.GetProperty("message").GetString());
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithBookmarks(int count)
    {
        var doc = new Document();
        for (var i = 0; i < count; i++)
            doc.Pages.Add();

        for (var i = 0; i < count; i++)
        {
            var bookmark = new OutlineItemCollection(doc.Outlines)
            {
                Title = $"Bookmark {i + 1}",
                Action = new GoToAction(doc.Pages[1])
            };
            doc.Outlines.Add(bookmark);
        }

        return doc;
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_GetsBookmarks()
    {
        var doc = CreateDocumentWithBookmarks(2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.True(json.RootElement.TryGetProperty("count", out _));
    }

    [Fact]
    public void Execute_ReturnsCorrectCount()
    {
        var doc = CreateDocumentWithBookmarks(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(3, json.RootElement.GetProperty("count").GetInt32());
    }

    [Fact]
    public void Execute_ReturnsItemsArray()
    {
        var doc = CreateDocumentWithBookmarks(2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(2, json.RootElement.GetProperty("items").GetArrayLength());
    }

    #endregion

    #region Bookmark Details

    [Fact]
    public void Execute_ReturnsBookmarkIndex()
    {
        var doc = CreateDocumentWithBookmarks(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);
        var firstBookmark = json.RootElement.GetProperty("items")[0];

        Assert.Equal(1, firstBookmark.GetProperty("index").GetInt32());
    }

    [Fact]
    public void Execute_ReturnsBookmarkTitle()
    {
        var doc = CreateDocumentWithPages(2);
        var bookmark = new OutlineItemCollection(doc.Outlines)
        {
            Title = "My Custom Title",
            Action = new GoToAction(doc.Pages[1])
        };
        doc.Outlines.Add(bookmark);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);
        var firstBookmark = json.RootElement.GetProperty("items")[0];

        Assert.Equal("My Custom Title", firstBookmark.GetProperty("title").GetString());
    }

    #endregion
}
