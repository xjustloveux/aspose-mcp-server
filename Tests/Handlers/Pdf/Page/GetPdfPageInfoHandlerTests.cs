using AsposeMcpServer.Handlers.Pdf.Page;
using AsposeMcpServer.Results.Pdf.Page;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Page;

public class GetPdfPageInfoHandlerTests : PdfHandlerTestBase
{
    private readonly GetPdfPageInfoHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetInfo()
    {
        Assert.Equal("get_info", _handler.Operation);
    }

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_DoesNotModifyDocument()
    {
        var doc = CreateDocumentWithPages(3);
        var initialCount = doc.Pages.Count;
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        Assert.Equal(initialCount, doc.Pages.Count);
        AssertNotModified(context);
    }

    #endregion

    #region Empty Document

    [Fact]
    public void Execute_EmptyDocument_ReturnsZeroCount()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPdfPageInfoResult>(res);

        Assert.Equal(1, result.Count);
    }

    #endregion

    #region Basic Info Retrieval

    [Fact]
    public void Execute_ReturnsPageInfo()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPdfPageInfoResult>(res);

        Assert.Equal(3, result.Count);
        AssertNotModified(context);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(3)]
    public void Execute_ReturnsCorrectPageCount(int pageCount)
    {
        var doc = CreateDocumentWithPages(pageCount);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPdfPageInfoResult>(res);

        Assert.Equal(pageCount, result.Count);
        AssertNotModified(context);
    }

    [SkippableTheory]
    [InlineData(5)]
    [InlineData(10)]
    public void Execute_ReturnsCorrectPageCount_HighPageCount(int pageCount)
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf, $"{pageCount} pages exceeds 4-page limit in evaluation mode");
        var doc = CreateDocumentWithPages(pageCount);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPdfPageInfoResult>(res);

        Assert.Equal(pageCount, result.Count);
        AssertNotModified(context);
    }

    #endregion

    #region Page Items

    [Fact]
    public void Execute_ReturnsItemsArray()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPdfPageInfoResult>(res);

        Assert.NotNull(result.Items);
        Assert.Equal(3, result.Items.Count);
    }

    [Fact]
    public void Execute_ItemsContainPageIndex()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPdfPageInfoResult>(res);

        var firstItem = result.Items[0];
        Assert.Equal(1, firstItem.PageIndex);
    }

    [Fact]
    public void Execute_ItemsContainDimensions()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPdfPageInfoResult>(res);

        var firstItem = result.Items[0];
        Assert.True(firstItem.Width >= 0);
        Assert.True(firstItem.Height >= 0);
    }

    [Fact]
    public void Execute_ItemsContainRotation()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPdfPageInfoResult>(res);

        var firstItem = result.Items[0];
        Assert.NotNull(firstItem.Rotation);
    }

    #endregion
}
