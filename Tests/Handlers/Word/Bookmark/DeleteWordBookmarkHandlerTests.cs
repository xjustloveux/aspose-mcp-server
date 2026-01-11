using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Bookmark;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Bookmark;

public class DeleteWordBookmarkHandlerTests : WordHandlerTestBase
{
    private readonly DeleteWordBookmarkHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Delete()
    {
        Assert.Equal("delete", _handler.Operation);
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

    #region Basic Delete Operations

    [Fact]
    public void Execute_DeletesBookmark()
    {
        var doc = CreateDocumentWithBookmarks(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "Bookmark0" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("deleted successfully", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsDeletedName()
    {
        var doc = CreateDocumentWithBookmarks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "Bookmark0" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Bookmark0", result);
    }

    [Fact]
    public void Execute_ReturnsRemainingCount()
    {
        var doc = CreateDocumentWithBookmarks(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "Bookmark0" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Remaining bookmarks", result);
    }

    #endregion

    #region Keep Text Option

    [Fact]
    public void Execute_WithKeepTextTrue_RemovesBookmarkKeepsText()
    {
        var doc = CreateDocumentWithBookmarks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "Bookmark0" },
            { "keepText", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Keep text: Yes", result);
        Assert.Null(doc.Range.Bookmarks["Bookmark0"]);
    }

    [Fact]
    public void Execute_WithKeepTextFalse_RemovesBookmarkAndText()
    {
        var doc = CreateDocumentWithBookmarks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "Bookmark0" },
            { "keepText", false }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Keep text: No", result);
    }

    [Fact]
    public void Execute_DefaultKeepTextIsTrue()
    {
        var doc = CreateDocumentWithBookmarks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "Bookmark0" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Keep text: Yes", result);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutName_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithBookmarks(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("name", ex.Message);
    }

    [Fact]
    public void Execute_WithNonExistentBookmark_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithBookmarks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "NonExistent" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("not found", ex.Message);
    }

    [Fact]
    public void Execute_NoBookmarks_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("No bookmarks here");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "AnyName" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("not found", ex.Message);
    }

    #endregion
}
