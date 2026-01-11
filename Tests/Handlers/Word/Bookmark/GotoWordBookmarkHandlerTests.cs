using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Bookmark;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Bookmark;

public class GotoWordBookmarkHandlerTests : WordHandlerTestBase
{
    private readonly GotoWordBookmarkHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Goto()
    {
        Assert.Equal("goto", _handler.Operation);
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

    #region Basic Goto Operations

    [Fact]
    public void Execute_GetsBookmarkLocation()
    {
        var doc = CreateDocumentWithBookmarks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "Bookmark0" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Bookmark location information", result);
    }

    [Fact]
    public void Execute_ReturnsBookmarkName()
    {
        var doc = CreateDocumentWithBookmarks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "Bookmark0" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Bookmark name: Bookmark0", result);
    }

    [Fact]
    public void Execute_ReturnsBookmarkText()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("TestBookmark");
        builder.Write("My bookmark text");
        builder.EndBookmark("TestBookmark");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "TestBookmark" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Bookmark text: My bookmark text", result);
    }

    [Fact]
    public void Execute_ReturnsParagraphIndex()
    {
        var doc = CreateDocumentWithBookmarks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "Bookmark0" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Paragraph index:", result);
    }

    [Fact]
    public void Execute_ReturnsRangeLength()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("TestBookmark");
        builder.Write("12345");
        builder.EndBookmark("TestBookmark");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "TestBookmark" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("5 characters", result);
    }

    [Fact]
    public void Execute_ReturnsParagraphContent()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Some paragraph text before");
        builder.StartBookmark("TestBookmark");
        builder.Write("bookmark content");
        builder.EndBookmark("TestBookmark");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "TestBookmark" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Paragraph content:", result);
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
