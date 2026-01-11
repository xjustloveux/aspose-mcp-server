using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.PivotTable;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"count\": 0", result);
        Assert.Contains("No pivot tables found", result);
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_WithPivotTable_ReturnsPivotTableInfo()
    {
        var workbook = CreateWorkbookWithPivotTable();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("count", result.ToLower());
        Assert.Contains("items", result.ToLower());
    }

    [Fact]
    public void Execute_ReturnsJsonFormat()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("{", result);
        Assert.Contains("}", result);
    }

    #endregion
}
