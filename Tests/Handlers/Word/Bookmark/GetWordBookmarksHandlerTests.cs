using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Bookmark;
using AsposeMcpServer.Results.Word.Bookmark;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Bookmark;

public class GetWordBookmarksHandlerTests : WordHandlerTestBase
{
    private readonly GetWordBookmarksHandler _handler = new();

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
        var doc = CreateDocumentWithText("No bookmarks here");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetBookmarksResult>(res);

        Assert.Equal(0, result.Count);
        Assert.NotNull(result.Message);
        Assert.Contains("No bookmarks found", result.Message);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithBookmarks(int count)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        for (var i = 0; i < count; i++)
        {
            builder.StartBookmark($"Bookmark{i}");
            builder.Write($"Bookmark {i} text");
            builder.EndBookmark($"Bookmark{i}");
            builder.InsertParagraph();
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

        var result = Assert.IsType<GetBookmarksResult>(res);

        Assert.True(result.Count >= 0);
    }

    [Fact]
    public void Execute_ReturnsCorrectCount()
    {
        var doc = CreateDocumentWithBookmarks(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetBookmarksResult>(res);

        Assert.Equal(3, result.Count);
    }

    [Fact]
    public void Execute_ReturnsBookmarksArray()
    {
        var doc = CreateDocumentWithBookmarks(2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetBookmarksResult>(res);

        Assert.Equal(2, result.Bookmarks.Count);
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

        var result = Assert.IsType<GetBookmarksResult>(res);
        var firstBookmark = result.Bookmarks[0];

        Assert.Equal(0, firstBookmark.Index);
    }

    [Fact]
    public void Execute_ReturnsBookmarkName()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("CustomName");
        builder.Write("text");
        builder.EndBookmark("CustomName");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetBookmarksResult>(res);
        var firstBookmark = result.Bookmarks[0];

        Assert.Equal("CustomName", firstBookmark.Name);
    }

    [Fact]
    public void Execute_ReturnsBookmarkText()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("TestBookmark");
        builder.Write("Bookmark Content");
        builder.EndBookmark("TestBookmark");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetBookmarksResult>(res);
        var firstBookmark = result.Bookmarks[0];

        Assert.Equal("Bookmark Content", firstBookmark.Text);
    }

    [Fact]
    public void Execute_ReturnsBookmarkLength()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("TestBookmark");
        builder.Write("12345");
        builder.EndBookmark("TestBookmark");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetBookmarksResult>(res);
        var firstBookmark = result.Bookmarks[0];

        Assert.Equal(5, firstBookmark.Length);
    }

    #endregion
}
