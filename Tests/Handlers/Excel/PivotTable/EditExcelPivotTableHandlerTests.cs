using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.PivotTable;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.PivotTable;

public class EditExcelPivotTableHandlerTests : ExcelHandlerTestBase
{
    private readonly EditExcelPivotTableHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Edit()
    {
        Assert.Equal("edit", _handler.Operation);
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

    #region Basic Edit Operations

    [Fact]
    public void Execute_EditsName()
    {
        var workbook = CreateWorkbookWithPivotTable();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pivotTableIndex", 0 },
            { "name", "NewPivotName" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("edited", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("name=NewPivotName", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithShowRowGrand_SetsRowGrand()
    {
        var workbook = CreateWorkbookWithPivotTable();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pivotTableIndex", 0 },
            { "showRowGrand", false }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("showRowGrand=False", result.Message);
    }

    [Fact]
    public void Execute_WithRefreshData_RefreshesPivotTable()
    {
        var workbook = CreateWorkbookWithPivotTable();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pivotTableIndex", 0 },
            { "refreshData", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("refreshed", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithNoChanges_ReturnsNoChanges()
    {
        var workbook = CreateWorkbookWithPivotTable();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pivotTableIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("no changes", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutPivotTableIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithPivotTable();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
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
