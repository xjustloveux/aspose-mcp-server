using AsposeMcpServer.Handlers.Pdf.Bookmark;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Bookmark;

public class AddPdfBookmarkHandlerTests : PdfHandlerTestBase
{
    private readonly AddPdfBookmarkHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_ReturnsTitle()
    {
        var doc = CreateDocumentWithPages(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "title", "My Bookmark" },
            { "pageIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("My Bookmark", result.Message);
    }

    [SkippableFact]
    public void Execute_ReturnsPageIndex()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "Exceeds evaluation mode page limits");
        var doc = CreateDocumentWithPages(5);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "title", "Chapter" },
            { "pageIndex", 3 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("page 3", result.Message);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void Execute_AddsBookmarkToVariousPages(int pageIndex)
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "title", $"Page {pageIndex}" },
            { "pageIndex", pageIndex }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Added bookmark", result.Message);
    }

    [Fact]
    public void Execute_AddsMultipleBookmarks()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);

        _handler.Execute(context, CreateParameters(new Dictionary<string, object?>
        {
            { "title", "First" },
            { "pageIndex", 1 }
        }));
        _handler.Execute(context, CreateParameters(new Dictionary<string, object?>
        {
            { "title", "Second" },
            { "pageIndex", 2 }
        }));

        Assert.Equal(2, doc.Outlines.Count);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutTitle_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("title", ex.Message);
    }

    [Fact]
    public void Execute_WithoutPageIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "title", "Bookmark" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("pageIndex", ex.Message);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void Execute_WithPageIndexLessThanOne_ThrowsArgumentException(int invalidIndex)
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "title", "Bookmark" },
            { "pageIndex", invalidIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("pageIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithPageIndexGreaterThanPageCount_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "title", "Bookmark" },
            { "pageIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("pageIndex", ex.Message);
    }

    #endregion
}
