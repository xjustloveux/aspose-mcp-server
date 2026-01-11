using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Bookmark;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Bookmark;

public class EditWordBookmarkHandlerTests : WordHandlerTestBase
{
    private readonly EditWordBookmarkHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Edit()
    {
        Assert.Equal("edit", _handler.Operation);
    }

    #endregion

    #region Edit Both

    [Fact]
    public void Execute_WithBothNameAndText_UpdatesBoth()
    {
        var doc = CreateDocumentWithBookmarks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "Bookmark0" },
            { "newName", "NewBookmark" },
            { "newText", "New content" }
        });

        _handler.Execute(context, parameters);

        var bookmark = doc.Range.Bookmarks["NewBookmark"];
        Assert.NotNull(bookmark);
        Assert.Equal("New content", bookmark.Text);
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

    #region Basic Edit Operations

    [Fact]
    public void Execute_EditsBookmark()
    {
        var doc = CreateDocumentWithBookmarks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "Bookmark0" },
            { "newText", "Updated content" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("edited successfully", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsOriginalName()
    {
        var doc = CreateDocumentWithBookmarks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "Bookmark0" },
            { "newName", "NewBookmarkName" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Original name: Bookmark0", result);
    }

    #endregion

    #region Edit Name

    [Fact]
    public void Execute_WithNewName_RenamesBookmark()
    {
        var doc = CreateDocumentWithBookmarks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "Bookmark0" },
            { "newName", "RenamedBookmark" }
        });

        _handler.Execute(context, parameters);

        Assert.NotNull(doc.Range.Bookmarks["RenamedBookmark"]);
    }

    [Fact]
    public void Execute_WithNewName_ReturnsNewName()
    {
        var doc = CreateDocumentWithBookmarks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "Bookmark0" },
            { "newName", "NewName" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("New name: NewName", result);
    }

    #endregion

    #region Edit Text

    [Fact]
    public void Execute_WithNewText_UpdatesContent()
    {
        var doc = CreateDocumentWithBookmarks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "Bookmark0" },
            { "newText", "Brand new text" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Brand new text", doc.Range.Bookmarks["Bookmark0"].Text);
    }

    [Fact]
    public void Execute_WithNewText_ReturnsNewContent()
    {
        var doc = CreateDocumentWithBookmarks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "Bookmark0" },
            { "newText", "Updated text" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("New content: Updated text", result);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutName_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithBookmarks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "newText", "Updated" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("name", ex.Message);
    }

    [Fact]
    public void Execute_WithoutNewNameOrNewText_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithBookmarks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "Bookmark0" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("newName or newText", ex.Message);
    }

    [Fact]
    public void Execute_WithNonExistentBookmark_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithBookmarks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "NonExistent" },
            { "newText", "Updated" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("not found", ex.Message);
    }

    [Fact]
    public void Execute_WithDuplicateNewName_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithBookmarks(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "Bookmark0" },
            { "newName", "Bookmark1" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("already exists", ex.Message);
    }

    #endregion
}
