using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.PivotTable;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.PivotTable;

public class RefreshExcelPivotTableHandlerTests : ExcelHandlerTestBase
{
    private readonly RefreshExcelPivotTableHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Refresh()
    {
        Assert.Equal("refresh", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithPivotTable()
    {
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["B3"].PutValue(200);

        var pivotTables = worksheet.PivotTables;
        pivotTables.Add($"={worksheet.Name}!A1:B3", "D1", "TestPivot");

        return workbook;
    }

    #endregion

    #region Basic Refresh Operations

    [Fact]
    public void Execute_RefreshesSpecificPivotTable()
    {
        var workbook = CreateWorkbookWithPivotTable();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pivotTableIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("refreshed", result.ToLower());
        Assert.Contains("1", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithoutIndex_RefreshesAllPivotTables()
    {
        var workbook = CreateWorkbookWithPivotTable();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("refreshed", result.ToLower());
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithNoPivotTables_ThrowsInvalidOperationException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        Assert.Throws<InvalidOperationException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithPivotTable();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pivotTableIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
