using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Bookmark;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Bookmark;

public class AddWordBookmarkHandlerTests : WordHandlerTestBase
{
    private readonly AddWordBookmarkHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsBookmark()
    {
        var doc = CreateDocumentWithText("Sample text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "TestBookmark" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Bookmark added successfully", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsBookmarkName()
    {
        var doc = CreateDocumentWithText("Sample text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "MyBookmark" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("MyBookmark", result);
    }

    [Fact]
    public void Execute_WithText_IncludesText()
    {
        var doc = CreateDocumentWithText("Sample text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "BookmarkWithText" },
            { "text", "Bookmark content here" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Bookmark text: Bookmark content here", result);
    }

    [Fact]
    public void Execute_WithParagraphIndex_InsertsAtPosition()
    {
        var doc = CreateDocumentWithParagraphs(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "PositionedBookmark" },
            { "paragraphIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("after paragraph #1", result);
    }

    [Fact]
    public void Execute_WithParagraphIndexMinusOne_InsertsAtBeginning()
    {
        var doc = CreateDocumentWithParagraphs(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "StartBookmark" },
            { "paragraphIndex", -1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("beginning of document", result);
    }

    [Fact]
    public void Execute_WithoutParagraphIndex_InsertsAtEnd()
    {
        var doc = CreateDocumentWithText("Sample text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "EndBookmark" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("end of document", result);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutName_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("name", ex.Message);
    }

    [Fact]
    public void Execute_WithEmptyName_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("name", ex.Message);
    }

    [Fact]
    public void Execute_WithDuplicateName_ThrowsInvalidOperationException()
    {
        var doc = CreateDocumentWithBookmarks(1);
        var existingName = doc.Range.Bookmarks[0].Name;
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", existingName }
        });

        var ex = Assert.Throws<InvalidOperationException>(() => _handler.Execute(context, parameters));
        Assert.Contains("already exists", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "InvalidBookmark" },
            { "paragraphIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithParagraphs(int count)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        for (var i = 0; i < count; i++) builder.Writeln($"Paragraph {i}");
        return doc;
    }

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
}
