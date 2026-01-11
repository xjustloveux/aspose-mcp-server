using System.Text.Json;
using AsposeMcpServer.Handlers.Excel.Sheet;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Sheet;

public class GetExcelSheetsHandlerTests : ExcelHandlerTestBase
{
    private readonly GetExcelSheetsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_DoesNotModifyWorkbook()
    {
        var workbook = CreateEmptyWorkbook();
        var initialCount = workbook.Worksheets.Count;
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        Assert.Equal(initialCount, workbook.Worksheets.Count);
        AssertNotModified(context);
    }

    #endregion

    #region Workbook Name

    [Fact]
    public void Execute_ReturnsWorkbookName()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("workbookName", out _));
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_ReturnsSheetInfo()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("count", out _));
        Assert.True(json.RootElement.TryGetProperty("items", out _));
        AssertNotModified(context);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void Execute_ReturnsCorrectCount(int sheetCount)
    {
        var workbook = CreateEmptyWorkbook();
        for (var i = 1; i < sheetCount; i++)
            workbook.Worksheets.Add($"Sheet{i + 1}");
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(sheetCount, json.RootElement.GetProperty("count").GetInt32());
    }

    #endregion

    #region Items Array

    [Fact]
    public void Execute_ReturnsItemsArray()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("items", out var items));
        Assert.Equal(JsonValueKind.Array, items.ValueKind);
        Assert.Equal(2, items.GetArrayLength());
    }

    [Fact]
    public void Execute_ItemsContainIndex()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var items = json.RootElement.GetProperty("items");
        Assert.True(items[0].TryGetProperty("index", out var index));
        Assert.Equal(0, index.GetInt32());
    }

    [Fact]
    public void Execute_ItemsContainName()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var items = json.RootElement.GetProperty("items");
        Assert.True(items[0].TryGetProperty("name", out var name));
        Assert.Equal("Sheet1", name.GetString());
    }

    [Fact]
    public void Execute_ItemsContainVisibility()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var items = json.RootElement.GetProperty("items");
        Assert.True(items[0].TryGetProperty("visibility", out var visibility));
        Assert.Equal("Visible", visibility.GetString());
    }

    [Fact]
    public void Execute_ReturnsHiddenVisibilityForHiddenSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("HiddenSheet");
        workbook.Worksheets["HiddenSheet"].IsVisible = false;
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var items = json.RootElement.GetProperty("items");
        var hiddenSheet = items.EnumerateArray().First(i => i.GetProperty("name").GetString() == "HiddenSheet");
        Assert.Equal("Hidden", hiddenSheet.GetProperty("visibility").GetString());
    }

    #endregion
}
