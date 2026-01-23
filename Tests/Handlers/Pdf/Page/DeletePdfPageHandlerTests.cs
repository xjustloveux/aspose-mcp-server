using AsposeMcpServer.Handlers.Pdf.Page;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Page;

public class DeletePdfPageHandlerTests : PdfHandlerTestBase
{
    private readonly DeletePdfPageHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Delete()
    {
        Assert.Equal("delete", _handler.Operation);
    }

    #endregion

    #region Error Handling - Missing Parameter

    [Fact]
    public void Execute_WithoutPageIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("pageIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Document State Preservation

    [Fact]
    public void Execute_PreservesOtherPages()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 2 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(2, doc.Pages.Count);
    }

    #endregion

    #region Basic Delete Operations

    [Fact]
    public void Execute_DeletesPageFromDocument()
    {
        var doc = CreateDocumentWithPages(3);
        var initialCount = doc.Pages.Count;
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 2 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Deleted", result.Message);
        Assert.Equal(initialCount - 1, doc.Pages.Count);
        AssertModified(context);
    }

    [SkippableTheory]
    [InlineData(3, 1, false)]
    [InlineData(3, 2, false)]
    [InlineData(3, 3, false)]
    [InlineData(5, 1, true)]
    [InlineData(5, 5, true)]
    public void Execute_DeletesPageAtVariousIndices(int totalPages, int deleteIndex, bool requiresLicense)
    {
        if (requiresLicense)
            SkipInEvaluationMode(AsposeLibraryType.Pdf, "Exceeds evaluation mode page limit");
        var doc = CreateDocumentWithPages(totalPages);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", deleteIndex }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Deleted", result.Message);
        Assert.Equal(totalPages - 1, doc.Pages.Count);
        AssertModified(context);
    }

    [Fact]
    public void Execute_DeletesFirstPage()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Deleted", result.Message);
        Assert.Equal(2, doc.Pages.Count);
        AssertModified(context);
    }

    [Fact]
    public void Execute_DeletesLastPage()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 3 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Deleted", result.Message);
        Assert.Equal(2, doc.Pages.Count);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_DeletesMiddlePage()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "Exceeds evaluation mode page limit");
        var doc = CreateDocumentWithPages(5);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 3 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Deleted", result.Message);
        Assert.Equal(4, doc.Pages.Count);
        AssertModified(context);
    }

    #endregion

    #region Result Message

    [SkippableFact]
    public void Execute_ReturnsRemainingPageCount()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "Exceeds evaluation mode page limit");
        var doc = CreateDocumentWithPages(5);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 2 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("4", result.Message);
    }

    [SkippableFact]
    public void Execute_ReturnsDeletedPageIndex()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "Exceeds evaluation mode page limit");
        var doc = CreateDocumentWithPages(5);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 3 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("3", result.Message);
    }

    #endregion

    #region Error Handling - Invalid Index

    [Theory]
    [InlineData(3, 4)]
    [InlineData(3, 5)]
    [InlineData(3, 10)]
    [InlineData(3, 100)]
    public void Execute_WithIndexOutOfRange_ThrowsArgumentException(int totalPages, int invalidIndex)
    {
        var doc = CreateDocumentWithPages(totalPages);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", invalidIndex }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithZeroPageIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(-5)]
    [InlineData(-100)]
    public void Execute_WithNegativePageIndex_ThrowsArgumentException(int negativeIndex)
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", negativeIndex }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Sequential Deletion

    [SkippableFact]
    public void Execute_SequentialDeletion_DeletesCorrectPages()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "Exceeds evaluation mode page limit");
        var doc = CreateDocumentWithPages(5);
        var context = CreateContext(doc);

        _handler.Execute(context, CreateParameters(new Dictionary<string, object?> { { "pageIndex", 1 } }));
        Assert.Equal(4, doc.Pages.Count);

        _handler.Execute(context, CreateParameters(new Dictionary<string, object?> { { "pageIndex", 1 } }));
        Assert.Equal(3, doc.Pages.Count);

        _handler.Execute(context, CreateParameters(new Dictionary<string, object?> { { "pageIndex", 1 } }));
        Assert.Equal(2, doc.Pages.Count);
    }

    [Fact]
    public void Execute_CanDeleteUntilOneRemains()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);

        _handler.Execute(context, CreateParameters(new Dictionary<string, object?> { { "pageIndex", 1 } }));
        Assert.Equal(2, doc.Pages.Count);

        _handler.Execute(context, CreateParameters(new Dictionary<string, object?> { { "pageIndex", 1 } }));
        Assert.Single(doc.Pages);
    }

    #endregion
}
