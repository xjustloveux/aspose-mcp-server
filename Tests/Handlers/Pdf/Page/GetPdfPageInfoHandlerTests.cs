using System.Text.Json;
using AsposeMcpServer.Handlers.Pdf.Page;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(1, json.RootElement.GetProperty("count").GetInt32());
    }

    #endregion

    #region Basic Info Retrieval

    [Fact]
    public void Execute_ReturnsPageInfo()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("count", out var count));
        Assert.Equal(3, count.GetInt32());
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

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(pageCount, json.RootElement.GetProperty("count").GetInt32());
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

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(pageCount, json.RootElement.GetProperty("count").GetInt32());
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

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("items", out var items));
        Assert.Equal(JsonValueKind.Array, items.ValueKind);
        Assert.Equal(3, items.GetArrayLength());
    }

    [Fact]
    public void Execute_ItemsContainPageIndex()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var items = json.RootElement.GetProperty("items");
        var firstItem = items[0];
        Assert.True(firstItem.TryGetProperty("pageIndex", out var pageIndex));
        Assert.Equal(1, pageIndex.GetInt32());
    }

    [Fact]
    public void Execute_ItemsContainDimensions()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var items = json.RootElement.GetProperty("items");
        var firstItem = items[0];
        Assert.True(firstItem.TryGetProperty("width", out _));
        Assert.True(firstItem.TryGetProperty("height", out _));
    }

    [Fact]
    public void Execute_ItemsContainRotation()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var items = json.RootElement.GetProperty("items");
        var firstItem = items[0];
        Assert.True(firstItem.TryGetProperty("rotation", out _));
    }

    #endregion
}
