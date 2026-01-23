using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.PivotTable;
using AsposeMcpServer.Results.Excel.PivotTable;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.PivotTable;

public class GetExcelPivotTablesHandlerTests : ExcelHandlerTestBase
{
    private readonly GetExcelPivotTablesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
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

    #region Basic Get Operations

    [Fact]
    public void Execute_WithNoPivotTables_ReturnsEmptyList()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPivotTablesResult>(res);

        Assert.Equal(0, result.Count);
        Assert.Equal("No pivot tables found", result.Message);
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_WithPivotTable_ReturnsPivotTableInfo()
    {
        var workbook = CreateWorkbookWithPivotTable();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPivotTablesResult>(res);

        Assert.True(result.Count > 0);
        Assert.NotNull(result.Items);
    }

    [Fact]
    public void Execute_ReturnsWorksheetName()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPivotTablesResult>(res);

        Assert.NotNull(result.WorksheetName);
    }

    #endregion
}
