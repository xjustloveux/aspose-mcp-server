using AsposeMcpServer.Handlers.Pdf.Page;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Page;

public class AddPdfPageHandlerTests : PdfHandlerTestBase
{
    private readonly AddPdfPageHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Count and InsertAt Combined

    [SkippableTheory]
    [InlineData(2, 1, false)]
    [InlineData(3, 2, true)]
    [InlineData(5, 1, true)]
    public void Execute_WithCountAndInsertAt_InsertsMultiplePagesAtPosition(int count, int insertAt,
        bool requiresLicense)
    {
        if (requiresLicense)
            SkipInEvaluationMode(AsposeLibraryType.Pdf, "Exceeds evaluation mode page limit");
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "count", count },
            { "insertAt", insertAt }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains(count.ToString(), result.Message);
        Assert.Equal(3 + count, doc.Pages.Count);
        AssertModified(context);
    }

    #endregion

    #region Basic Page Addition

    [Fact]
    public void Execute_AddsPageToDocument()
    {
        var doc = CreateEmptyDocument();
        var initialCount = doc.Pages.Count;
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Added", result.Message);
        Assert.Equal(initialCount + 1, doc.Pages.Count);
        AssertModified(context);
    }

    [Fact]
    public void Execute_MultipleTimes_AddsMultiplePages()
    {
        var doc = CreateEmptyDocument();
        var initialCount = doc.Pages.Count;
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);
        _handler.Execute(context, parameters);
        _handler.Execute(context, parameters);

        Assert.Equal(initialCount + 3, doc.Pages.Count);
    }

    [Fact]
    public void Execute_DefaultCount_AddsOnePage()
    {
        var doc = CreateEmptyDocument();
        var initialCount = doc.Pages.Count;
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        Assert.Equal(initialCount + 1, doc.Pages.Count);
    }

    #endregion

    #region Count Parameter

    [SkippableTheory]
    [InlineData(1, false)]
    [InlineData(2, false)]
    [InlineData(3, false)]
    [InlineData(5, true)]
    [InlineData(10, true)]
    public void Execute_WithCount_AddsCorrectNumberOfPages(int count, bool requiresLicense)
    {
        if (requiresLicense)
            SkipInEvaluationMode(AsposeLibraryType.Pdf, "Exceeds evaluation mode page limit");
        var doc = CreateEmptyDocument();
        var initialCount = doc.Pages.Count;
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "count", count }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains(count.ToString(), result.Message);
        Assert.Equal(initialCount + count, doc.Pages.Count);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithZeroCount_AddsDefaultOnePage()
    {
        var doc = CreateEmptyDocument();
        var initialCount = doc.Pages.Count;
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "count", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.True(doc.Pages.Count >= initialCount);
    }

    #endregion

    #region InsertAt Parameter

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void Execute_WithInsertAt_InsertsAtCorrectPosition(int insertAt)
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "insertAt", insertAt }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Added", result.Message);
        Assert.Equal(4, doc.Pages.Count);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithInsertAtBeyondEnd_AppendsToEnd()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "insertAt", 100 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Added", result.Message);
        Assert.Equal(4, doc.Pages.Count);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithInsertAtZero_InsertsAtBeginning()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "insertAt", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Added", result.Message);
        Assert.Equal(4, doc.Pages.Count);
        AssertModified(context);
    }

    #endregion

    #region Page Size

    [Theory]
    [InlineData(612.0, 792.0)]
    [InlineData(595.0, 842.0)]
    [InlineData(841.0, 1190.0)]
    public void Execute_WithCustomSize_SetsCorrectPageSize(double width, double height)
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "width", width },
            { "height", height }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Added", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithOnlyWidth_UsesDefaultHeight()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "width", 612.0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Added", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithOnlyHeight_UsesDefaultWidth()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "height", 792.0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Added", result.Message);
        AssertModified(context);
    }

    #endregion

    #region Result Message

    [SkippableTheory]
    [InlineData(1, 1, false)]
    [InlineData(3, 2, true)]
    [InlineData(5, 3, true)]
    public void Execute_ReturnsTotalPageCount(int initialPages, int addedPages, bool requiresLicense)
    {
        if (requiresLicense)
            SkipInEvaluationMode(AsposeLibraryType.Pdf, "Exceeds evaluation mode page limit");
        var doc = CreateDocumentWithPages(initialPages);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "count", addedPages }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        var expectedTotal = initialPages + addedPages;
        Assert.Contains(expectedTotal.ToString(), result.Message);
    }

    [Fact]
    public void Execute_ReturnsAddedPageCount()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "count", 3 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("3", result.Message);
        Assert.Contains("Added", result.Message);
    }

    #endregion

    #region Document State Preservation

    [Fact]
    public void Execute_PreservesExistingPages()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        Assert.Equal(4, doc.Pages.Count);
    }

    [Fact]
    public void Execute_AddsPageAtEnd_ByDefault()
    {
        var doc = CreateDocumentWithPages(3);
        var initialPageCount = doc.Pages.Count;
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        Assert.Equal(initialPageCount + 1, doc.Pages.Count);
    }

    #endregion
}
