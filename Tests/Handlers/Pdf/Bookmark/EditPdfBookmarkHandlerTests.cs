using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Handlers.Pdf.Bookmark;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Bookmark;

public class EditPdfBookmarkHandlerTests : PdfHandlerTestBase
{
    private readonly EditPdfBookmarkHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Edit()
    {
        Assert.Equal("edit", _handler.Operation);
    }

    #endregion

    #region Edit Title

    [Fact]
    public void Execute_WithTitle_ChangesTitle()
    {
        var doc = CreateDocumentWithBookmarks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "bookmarkIndex", 1 },
            { "title", "Brand New Title" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Brand New Title", doc.Outlines[1].Title);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithBookmarks(int count)
    {
        var doc = new Document();
        for (var i = 0; i < 3; i++)
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

    #region Basic Edit Operations

    [Fact]
    public void Execute_EditsBookmark()
    {
        var doc = CreateDocumentWithBookmarks(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "bookmarkIndex", 1 },
            { "title", "New Title" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Edited bookmark", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsBookmarkIndex()
    {
        var doc = CreateDocumentWithBookmarks(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "bookmarkIndex", 1 },
            { "title", "Updated" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("index 1", result.Message);
    }

    #endregion

    #region Edit Page Index

    [Fact]
    public void Execute_WithPageIndex_ChangesPageIndex()
    {
        var doc = CreateDocumentWithBookmarks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "bookmarkIndex", 1 },
            { "pageIndex", 2 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Edited bookmark", result.Message);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void Execute_WithVariousPageIndices_Works(int pageIndex)
    {
        var doc = CreateDocumentWithBookmarks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "bookmarkIndex", 1 },
            { "pageIndex", pageIndex }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Edited bookmark", result.Message);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutBookmarkIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithBookmarks(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "title", "New Title" }
        });

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
            { "bookmarkIndex", invalidIndex },
            { "title", "New Title" }
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
            { "bookmarkIndex", 99 },
            { "title", "New Title" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("bookmarkIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidPageIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithBookmarks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "bookmarkIndex", 1 },
            { "pageIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("pageIndex", ex.Message);
    }

    #endregion
}
