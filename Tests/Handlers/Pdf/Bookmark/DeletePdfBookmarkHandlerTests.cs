using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Handlers.Pdf.Bookmark;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Bookmark;

public class DeletePdfBookmarkHandlerTests : PdfHandlerTestBase
{
    private readonly DeletePdfBookmarkHandler _handler = new();

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

    #region Basic Delete Operations

    [Fact]
    public void Execute_DeletesBookmark()
    {
        var doc = CreateDocumentWithBookmarks(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "bookmarkIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Deleted bookmark", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsDeletedTitle()
    {
        var doc = CreateDocumentWithPages(2);
        var bookmark = new OutlineItemCollection(doc.Outlines)
        {
            Title = "ToDelete",
            Action = new GoToAction(doc.Pages[1])
        };
        doc.Outlines.Add(bookmark);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "bookmarkIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("ToDelete", result);
    }

    [Fact]
    public void Execute_ReturnsDeletedIndex()
    {
        var doc = CreateDocumentWithBookmarks(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "bookmarkIndex", 2 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("index 2", result);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void Execute_DeletesAtVariousIndices(int index)
    {
        var doc = CreateDocumentWithBookmarks(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "bookmarkIndex", index }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Deleted bookmark", result);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutBookmarkIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithBookmarks(2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("bookmarkIndex", ex.Message);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void Execute_WithBookmarkIndexLessThanOne_ThrowsArgumentException(int invalidIndex)
    {
        var doc = CreateDocumentWithBookmarks(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "bookmarkIndex", invalidIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("bookmarkIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithBookmarkIndexGreaterThanCount_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithBookmarks(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "bookmarkIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("bookmarkIndex", ex.Message);
    }

    [Fact]
    public void Execute_NoBookmarks_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "bookmarkIndex", 1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("bookmarkIndex", ex.Message);
    }

    #endregion
}
