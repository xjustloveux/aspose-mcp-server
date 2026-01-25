using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Handlers.Excel.Sheet;
using AsposeMcpServer.Results.Excel.Sheet;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSheetsResult>(res);

        Assert.NotNull(result.WorkbookName);
    }

    [Fact]
    public void Execute_WithSourcePath_ReturnsFileName()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContextWithSourcePath(workbook, "/path/to/test_workbook.xlsx");
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSheetsResult>(res);

        Assert.Equal("test_workbook.xlsx", result.WorkbookName);
    }

    [Fact]
    public void Execute_WithoutSourcePath_ReturnsSession()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSheetsResult>(res);

        Assert.Equal("session", result.WorkbookName);
    }

    private static OperationContext<Workbook> CreateContextWithSourcePath(Workbook workbook, string sourcePath)
    {
        return new OperationContext<Workbook>
        {
            Document = workbook,
            SourcePath = sourcePath
        };
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_ReturnsSheetInfo()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSheetsResult>(res);

        Assert.True(result.Count >= 0);
        Assert.NotNull(result.Items);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSheetsResult>(res);

        Assert.Equal(sheetCount, result.Count);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSheetsResult>(res);

        Assert.NotNull(result.Items);
        Assert.Equal(2, result.Items.Count);
    }

    [Fact]
    public void Execute_ItemsContainIndex()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSheetsResult>(res);

        Assert.Equal(0, result.Items[0].Index);
    }

    [Fact]
    public void Execute_ItemsContainName()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSheetsResult>(res);

        Assert.Equal("Sheet1", result.Items[0].Name);
    }

    [Fact]
    public void Execute_ItemsContainVisibility()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSheetsResult>(res);

        Assert.Equal("Visible", result.Items[0].Visibility);
    }

    [Fact]
    public void Execute_ReturnsHiddenVisibilityForHiddenSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("HiddenSheet");
        workbook.Worksheets["HiddenSheet"].IsVisible = false;
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSheetsResult>(res);

        var hiddenSheet = result.Items.First(i => i.Name == "HiddenSheet");
        Assert.Equal("Hidden", hiddenSheet.Visibility);
    }

    #endregion
}
