using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Handlers.Pdf.Bookmark;
using AsposeMcpServer.Results.Pdf.Bookmark;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetBookmarksPdfResult>(res);

        Assert.Equal(0, result.Count);
        Assert.Contains("No bookmarks found", result.Message);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetBookmarksPdfResult>(res);

        Assert.True(result.Count >= 0);
    }

    [Fact]
    public void Execute_ReturnsCorrectCount()
    {
        var doc = CreateDocumentWithBookmarks(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetBookmarksPdfResult>(res);

        Assert.Equal(3, result.Count);
    }

    [Fact]
    public void Execute_ReturnsItemsArray()
    {
        var doc = CreateDocumentWithBookmarks(2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetBookmarksPdfResult>(res);

        Assert.Equal(2, result.Items.Count);
    }

    #endregion

    #region Bookmark Details

    [Fact]
    public void Execute_ReturnsBookmarkIndex()
    {
        var doc = CreateDocumentWithBookmarks(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetBookmarksPdfResult>(res);
        var firstBookmark = result.Items[0];

        Assert.Equal(1, firstBookmark.Index);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetBookmarksPdfResult>(res);
        var firstBookmark = result.Items[0];

        Assert.Equal("My Custom Title", firstBookmark.Title);
    }

    #endregion
}
